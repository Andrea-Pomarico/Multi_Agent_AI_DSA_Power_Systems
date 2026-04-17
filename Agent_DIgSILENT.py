"""
╔══════════════════════════════════════════════════════════════════╗
║           DIGSILENT AGENT — Standalone RMS Simulation            ║
║  Capabilities:                                                   ║
║    1. Activate project & study case                              ║
║    2. Run load flow (ComLdf)                                     ║
║    3. Export grid graph with line loading / flow                  ║
║    4. Run RMS simulation (ComInc + ComSim)                       ║
║    5. Export results to CSV (V, rotor angle, frequency)          ║
╚══════════════════════════════════════════════════════════════════╝
"""

import sys
import os
import json
import math
import re
import time
from dataclasses import dataclass, field
from typing import Optional

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch

# ── PowerFactory Python path ──────────────────────────────────────
sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP1\Python\3.13")

import powerfactory as pf

# ══════════════════════════════════════════════════════════════════
# CONFIGURATION — edit this block to match your setup
# ══════════════════════════════════════════════════════════════════

@dataclass
class SimulationConfig:
    """All parameters needed to run one RMS simulation."""

    # ── Project ───────────────────────────────────────────────────
    project_path: str = r"\ilea_\Andrea Pomarico\UW\39 Bus SINDy RMS MA.IntPrj"
    study_case:   str = r"ctocto"
    base_study_case: str = r"0. Base"

    # ── Fault ─────────────────────────────────────────────────────
    # fault_type : "bus"  → EvtShc ON + EvtShc OFF (clear)
    #              "line" → EvtShc ON + EvtSwitch OPEN (trip line)
    #              "gen_switch" → EvtSwitch on generator (open/close)
    fault_type:    str = "bus"
    fault_element: str = "Bus 01.ElmTerm"   # PF object name for the short-circuit
    switch_element: str = ""                # PF object name for generator switch event (e.g., Gen 05.ElmSym)
    t_switch: float = 1.0                    # time when generator switch is applied
    switch_state: int = 0                    # EvtSwitch.i_switch (0=open, 1=close)

    # ── RMS simulation timing (seconds) ──────────────────────────
    t_start: float = 0.0
    t_fault: float = 1.0    # time when fault is applied
    t_clear: float = 1.08   # fault clearance time  (FCT = 80 ms)
    t_end:   float = 10.0   # total simulation duration

    # ── Time step ─────────────────────────────────────────────────
    dt_rms: float = 0.01    # seconds

    # ── CSV output ────────────────────────────────────────────────
    output_dir:   str = r"C:\RMS_Results"
    run_label:    str = "run_001"
    result_name:  str = "All calculations.ElmRes"
    word_document: int = 0
    final_word_document: int = 1
    final_presentation: int = 1

    # Set to 1 to enable optional LLM pipeline steps.
    # Disabled by default to reduce API quota usage on quick test runs.
    run_review_agent: int = 0
    run_final_report_agent: int = 0
    run_mitigation_agent: int = 0

    # ──────────────────────────────────────────────────────────────
    @classmethod
    def from_json(cls, path: str) -> "SimulationConfig":
        """Load config from a JSON file, overriding only the keys present."""
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return cls(**{k: v for k, v in data.items()
                      if k in cls.__dataclass_fields__})

    # ── Signals to export ─────────────────────────────────────────
    # Each entry: (object_name, variable_name, friendly_label)
    # Adjust names to match elements in your network model.
    signals: list = field(default_factory=lambda: [
        # Bus voltages
        ("Bus 01.ElmTerm",    "m:u",    "V_Troia_pu"),
        ("Bus 02.ElmTerm",   "m:u",    "V_Ariano_pu"),
        ("Bus 03.ElmTerm",   "m:u",    "V_Latina_pu"),
        ("Bus 04.ElmTerm","m:u",   "V_Garigliano_pu"),

        # # Generator rotor angles
        # ("Gen 01.ElmSym",       "s:firel","Angle_CS1_deg"),
        # ("Gen 02.ElmSym",       "s:firel","Angle_CS2_deg"),

        # # System frequency (measured at reference bus or machine)
        # ("Gen 01.ElmSym",       "m:f",    "Freq_CS1_Hz"),
    ])


# ══════════════════════════════════════════════════════════════════
# LOGGER
# ══════════════════════════════════════════════════════════════════

class Logger:
    """Simple timestamped console logger."""

    @staticmethod
    def info(msg: str):  print(f"[INFO]  {time.strftime('%H:%M:%S')} | {msg}")

    @staticmethod
    def ok(msg: str):    print(f"[OK]    {time.strftime('%H:%M:%S')} | ✅ {msg}")

    @staticmethod
    def warn(msg: str):  print(f"[WARN]  {time.strftime('%H:%M:%S')} | ⚠️  {msg}")

    @staticmethod
    def error(msg: str): print(f"[ERROR] {time.strftime('%H:%M:%S')} | ❌ {msg}")

    @staticmethod
    def section(title: str):
        bar = "═" * 60
        print(f"\n{bar}\n  {title}\n{bar}")


log = Logger()


# ══════════════════════════════════════════════════════════════════
# DIGSILENT AGENT
# ══════════════════════════════════════════════════════════════════

class DIgSILENTAgent:
    """
    Standalone agent that wraps the PowerFactory Python API.
    All public methods return (success: bool, message: str).
    """

    # Keep one PowerFactory handle per Python process.
    # PowerFactory cannot be started multiple times in the same process.
    _shared_app: Optional[object] = None
    _shared_project_path: Optional[str] = None
    _shared_project: Optional[object] = None

    def __init__(self, config: SimulationConfig):
        self.cfg = config
        self.app: Optional[object] = None
        self.project: Optional[object] = None
        self.result_objects: dict = {}   # label → PF result object
        # Populated by export_grid_graph; used by downstream agents.
        self.grid_data: dict = {}

    # ──────────────────────────────────────────────────────────────
    # STEP 1 — Connect to PowerFactory & activate project
    # ──────────────────────────────────────────────────────────────

    def connect(self) -> tuple[bool, str]:
        log.section("STEP 1 — Connect to PowerFactory")
        try:
            if DIgSILENTAgent._shared_app is None:
                self.app = pf.GetApplicationExt()
                if self.app is None:
                    raise RuntimeError("GetApplicationExt() returned None")
                self.app.Show()
                DIgSILENTAgent._shared_app = self.app
                log.ok("PowerFactory application obtained and shown")
            else:
                self.app = DIgSILENTAgent._shared_app
                log.ok("Reusing existing PowerFactory application in this process")
        except Exception as e:
            log.error(f"Cannot connect to PowerFactory: {e}")
            return False, str(e)

        try:
            if DIgSILENTAgent._shared_project_path != self.cfg.project_path:
                self.project = self.app.ActivateProject(self.cfg.project_path)
                if self.project is None:
                    raise RuntimeError(f"Project not found: {self.cfg.project_path}")
                DIgSILENTAgent._shared_project = self.project
                DIgSILENTAgent._shared_project_path = self.cfg.project_path
                log.ok(f"Project activated: {self.cfg.project_path}")
            else:
                self.project = DIgSILENTAgent._shared_project
                log.ok(f"Reusing already active project: {self.cfg.project_path}")
        except Exception as e:
            log.error(f"Cannot activate project: {e}")
            return False, str(e)

        return True, "Connected and project activated"

    # ──────────────────────────────────────────────────────────────
    # STEP 2 — Activate study case
    # ──────────────────────────────────────────────────────────────

    def activate_study_case(self) -> tuple[bool, str]:
        log.section("STEP 2 — Activate Study Case")
        try:
            folder = self.app.GetProjectFolder('study')
            target_name = self.cfg.study_case
            base_name = getattr(self.cfg, "base_study_case", "0. Base")

            if folder is not None:
                # Standard project: study cases folder exists
                target_contents = folder.GetContents(target_name)
                if target_contents:
                    target_contents[0].Activate()
                else:
                    base_contents = folder.GetContents(base_name)
                    if not base_contents:
                        raise RuntimeError(
                            f"Base study case not found in study folder: '{base_name}'"
                        )
                    if target_name == base_name:
                        base_contents[0].Activate()
                    else:
                        new_study_case = folder.AddCopy(base_contents[0], target_name)
                        if new_study_case is None:
                            target_contents = folder.GetContents(target_name)
                            if not target_contents:
                                raise RuntimeError(
                                    f"Study case copy failed: '{target_name}'"
                                )
                            new_study_case = target_contents[0]
                        new_study_case.Activate()
                        log.ok(f"Study case copied from '{base_name}' to '{target_name}'")
            else:
                # Non-standard project: search the whole project for *.IntCase by name
                log.warn("GetProjectFolder('study') returned None — searching project for IntCase objects")
                case_name = target_name.split('\\')[-1]
                matches = self.app.GetCalcRelevantObjects(f"{case_name}.IntCase")
                if not matches:
                    raise RuntimeError(
                        f"Study case '{case_name}' not found via GetCalcRelevantObjects either. "
                        "Check the name in PowerFactory's Data Manager."
                    )
                matches[0].Activate()

            log.ok(f"Study case activated: {target_name}")
            return True, "Study case activated"
        except Exception as e:
            log.error(f"Study case activation failed: {e}")
            return False, str(e)

    @staticmethod
    def _line_object_name(obj) -> str | None:
        if obj is None:
            return None
        if isinstance(obj, str):
            text = obj.strip()
            return text or None
        for attr in ("loc_name", "name"):
            value = getattr(obj, attr, None)
            if isinstance(value, str) and value.strip():
                return value.strip()
        try:
            text = str(obj).strip()
        except Exception:
            return None
        if not text or text.startswith("<"):
            return None
        return text

    def _line_terminal_name(self, line, attribute_names: tuple[str, ...]) -> str | None:
        for attribute_name in attribute_names:
            try:
                value = line.GetAttribute(attribute_name)
            except Exception:
                value = None
            name = self._line_object_name(value)
            if name:
                return name
        return None

    @staticmethod
    def _normalize_bus_token(token: str) -> str:
        clean = token.strip()
        if clean.isdigit():
            return str(int(clean))
        return clean

    @staticmethod
    def _buses_from_line_name(line_name: str | None) -> tuple[str | None, str | None]:
        if not line_name:
            return None, None
        # Expected patterns:
        #   "Line 01 - 02" -> "Bus 1", "Bus 2"
        #   "Trf 02 - 30"  -> "Bus 2", "Bus 30"
        match = re.search(r"(?i)\b(?:line|trf)\s*([0-9A-Za-z]+)\s*-\s*([0-9A-Za-z]+)\b", line_name)
        if not match:
            return None, None
        left = DIgSILENTAgent._normalize_bus_token(match.group(1))
        right = DIgSILENTAgent._normalize_bus_token(match.group(2))
        if not left or not right:
            return None, None
        return f"Bus {left}", f"Bus {right}"

    @staticmethod
    def _bus_from_load_name(load_name: str | None) -> str | None:
        if not load_name:
            return None
        match = re.search(r"(?i)\bload\s*([0-9A-Za-z]+)\b", load_name)
        if not match:
            return None
        bus_token = DIgSILENTAgent._normalize_bus_token(match.group(1))
        if not bus_token:
            return None
        return f"Bus {bus_token}"

    @staticmethod
    def _bus_from_generator_name(gen_name: str | None) -> str | None:
        if not gen_name:
            return None

        number_map = {
            "01": "Bus 39",
            "1": "Bus 39",
            "02": "Bus 31",
            "2": "Bus 31",
            "03": "Bus 32",
            "3": "Bus 32",
            "04": "Bus 33",
            "4": "Bus 33",
            "05": "Bus 34",
            "5": "Bus 34",
            "06": "Bus 35",
            "6": "Bus 35",
            "07": "Bus 36",
            "7": "Bus 36",
            "08": "Bus 37",
            "8": "Bus 37",
            "09": "Bus 38",
            "9": "Bus 38",
            "10": "Bus 30",
        }

        normalized = re.sub(r"\s+", " ", gen_name.strip().lower())
        normalized = normalized.split(".", 1)[0].strip()
        match = re.search(r"(\d+)", normalized)
        if match:
            return number_map.get(match.group(1).zfill(2)) or number_map.get(match.group(1))
        return None

    @staticmethod
    def _as_float(value) -> float | None:
        try:
            if value is None:
                return None
            return float(value)
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _bus_key(bus_name: str | None) -> str | None:
        if not bus_name:
            return None
        match = re.search(r"(?i)\bbus\s*([0-9A-Za-z]+)\b", bus_name)
        if not match:
            return bus_name.strip().lower() or None
        return f"bus_{DIgSILENTAgent._normalize_bus_token(match.group(1)).lower()}"

    def _bus_characteristics(self) -> dict[str, dict[str, float | None]]:
        bus_data: dict[str, dict[str, float | None]] = {}
        terminals = self.app.GetCalcRelevantObjects("*.ElmTerm") or []

        # Capture load-flow quantities first.
        for terminal in terminals:
            terminal_name = self._line_object_name(getattr(terminal, "loc_name", None))
            if not terminal_name:
                continue
            key = self._bus_key(terminal_name)
            if not key:
                continue
            bus_data[key] = {
                "name": terminal_name,
                "u": self._as_float(terminal.GetAttribute("m:u")),
                "phiu": self._as_float(terminal.GetAttribute("m:phiu")),
                "ikss": None,
            }

        # Then run short-circuit and append Ikss.
        try:
            shc = self.app.GetFromStudyCase("ComShc")
            if shc is not None:
                err = shc.Execute()
                if err:
                    log.warn(f"ComShc returned error code {err}; Ikss may be unavailable")
            else:
                log.warn("ComShc object not found in study case; Ikss may be unavailable")
        except Exception as e:
            log.warn(f"Short-circuit calculation failed before graph export: {e}")

        for terminal in terminals:
            terminal_name = self._line_object_name(getattr(terminal, "loc_name", None))
            if not terminal_name:
                continue
            key = self._bus_key(terminal_name)
            if not key:
                continue
            if key not in bus_data:
                bus_data[key] = {"name": terminal_name, "u": None, "phiu": None, "ikss": None}
            bus_data[key]["ikss"] = self._as_float(terminal.GetAttribute("m:Ikss"))
        return bus_data

    def export_grid_graph(self) -> tuple[bool, str]:
        log.section("STEP 3 — Export Grid Graph")
        try:
            os.makedirs(self.cfg.output_dir, exist_ok=True)

            lines = self.app.GetCalcRelevantObjects("*.ElmLne") or []
            trf2_list = self.app.GetCalcRelevantObjects("*.ElmTr2") or []
            load_list = self.app.GetCalcRelevantObjects("*.ElmLod") or []
            gen_list = self.app.GetCalcRelevantObjects("*.ElmSym") or []
            if not lines and not trf2_list and not load_list and not gen_list:
                raise RuntimeError("No ElmLne, ElmTr2, ElmLod or ElmSym objects found in the active study case")

            line_data = []
            load_data = []
            generator_data = []
            bus_names = set()
            skipped = 0

            def _append_edge(element, element_type: str):
                nonlocal skipped
                line_name = self._line_object_name(element.GetAttribute("b:loc_name"))
                if not line_name:
                    line_name = self._line_object_name(getattr(element, "loc_name", None))
                if not line_name:
                    line_name = f"Line_{len(line_data) + skipped + 1}"

                bus1, bus2 = self._buses_from_line_name(line_name)
                if not bus1 or not bus2:
                    if element_type == "line":
                        bus1 = self._line_terminal_name(element, ("e:bus1_bar", "bus1", "b:bus1"))
                        bus2 = self._line_terminal_name(element, ("e:bus2_bar", "bus2", "b:bus2"))
                    else:
                        bus1 = self._line_terminal_name(element, ("e:bushv_bar", "bushv", "b:bushv"))
                        bus2 = self._line_terminal_name(element, ("e:buslv_bar", "buslv", "b:buslv"))

                loading = self._as_float(element.GetAttribute("c:loading"))
                if element_type == "line":
                    p_flow = self._as_float(element.GetAttribute("n:Pflow:bus1"))
                    if p_flow is None:
                        p_flow = self._as_float(element.GetAttribute("m:P:bus1"))
                else:
                    p_flow = self._as_float(element.GetAttribute("m:P:bushv"))
                # print(f"  processing {element_type} '{line_name}': Loading: {loading}, Pflow: {p_flow}")

                if not bus1 or not bus2:
                    skipped += 1
                    return

                bus_names.update((bus1, bus2))
                line_data.append(
                    {
                        "name": line_name,
                        "bus1": bus1,
                        "bus2": bus2,
                        "loading": loading,
                        "p_flow": p_flow,
                    }
                )

            for line in lines:
                _append_edge(line, "line")
            for trf in trf2_list:
                _append_edge(trf, "trf2")

            def _append_load(load):
                nonlocal skipped
                load_name = self._line_object_name(load.GetAttribute("b:loc_name"))
                if not load_name:
                    load_name = self._line_object_name(getattr(load, "loc_name", None))
                if not load_name:
                    load_name = f"Load_{len(load_data) + skipped + 1}"

                bus1 = self._bus_from_load_name(load_name)
                if not bus1:
                    bus1 = self._line_terminal_name(load, ("e:bus1_bar", "bus1", "b:bus1"))

                p_ini = self._as_float(load.GetAttribute("c:pini"))
                q_ini = self._as_float(load.GetAttribute("c:qini"))
                # print(f"  processing load '{load_name}': P={p_ini}, Q={q_ini}")

                if not bus1:
                    skipped += 1
                    return

                bus_names.add(bus1)
                load_data.append(
                    {
                        "name": load_name,
                        "bus": bus1,
                        "p_ini": p_ini,
                        "q_ini": q_ini,
                    }
                )

            for load in load_list:
                _append_load(load)

            def _append_generator(generator):
                nonlocal skipped
                gen_name = self._line_object_name(generator.GetAttribute("b:loc_name"))
                if not gen_name:
                    gen_name = self._line_object_name(getattr(generator, "loc_name", None))
                if not gen_name:
                    gen_name = f"Gen_{len(generator_data) + skipped + 1}"

                bus_name = self._bus_from_generator_name(gen_name)
                if not bus_name:
                    bus_name = self._line_terminal_name(generator, ("e:bus1_bar", "bus1", "b:bus1"))

                p_ini = self._as_float(generator.GetAttribute("m:P:bus1"))
                q_ini = self._as_float(generator.GetAttribute("m:Q:bus1"))
                # print(f"  processing generator '{gen_name}': P={p_ini}, Q={q_ini}")

                if not bus_name:
                    skipped += 1
                    return

                bus_names.add(bus_name)
                generator_data.append(
                    {
                        "name": gen_name,
                        "bus": bus_name,
                        "p_ini": p_ini,
                        "q_ini": q_ini,
                    }
                )

            for generator in gen_list:
                _append_generator(generator)

            if not line_data and not load_data and not generator_data:
                raise RuntimeError("Could not resolve any edges, loads, or generators for the graph")

            try:
                import networkx as nx
            except Exception:
                nx = None

            safe_label = re.sub(r"[^A-Za-z0-9_.-]+", "_", self.cfg.run_label)
            graph_path = os.path.join(self.cfg.output_dir, f"{safe_label}_grid_graph.png")
            bus_characteristics = self._bus_characteristics()

            def _bus_label(bus_name: str) -> str:
                info = bus_characteristics.get(self._bus_key(bus_name), {})
                voltage = info.get("u") if isinstance(info, dict) else None
                angle = info.get("phiu") if isinstance(info, dict) else None
                ikss = info.get("ikss") if isinstance(info, dict) else None
                if voltage is None and angle is None and ikss is None:
                    return bus_name
                voltage_txt = "n/a" if voltage is None else f"{voltage:.3f}"
                angle_txt = "n/a" if angle is None else f"{angle:.2f}"
                ikss_txt = "n/a" if ikss is None else f"{ikss:.2f}"
                return f"{bus_name}\nV={voltage_txt} pu\nang={angle_txt} deg\nIkss={ikss_txt} kA"

            if nx is not None:
                graph = nx.MultiGraph()
                for bus_name in sorted(bus_names):
                    graph.add_node(bus_name)
                for item in line_data:
                    graph.add_edge(
                        item["bus1"],
                        item["bus2"],
                        key=item["name"],
                        name=item["name"],
                        loading=item["loading"],
                        p_flow=item["p_flow"],
                    )

                pos = nx.spring_layout(graph, seed=42, k=1.1, iterations=200)
                fig, (ax, ax_info) = plt.subplots(
                    1,
                    2,
                    figsize=(22, 12),
                    gridspec_kw={"width_ratios": [4.6, 1.4]},
                )
                ax.set_title(f"PowerFactory Grid Topology — {self.cfg.study_case}")
                ax.axis("off")
                ax_info.axis("off")

                nx.draw_networkx_nodes(
                    graph,
                    pos,
                    ax=ax,
                    node_size=1100,
                    node_color="#f7f2df",
                    edgecolors="#333333",
                    linewidths=0.9,
                )
                nx.draw_networkx_labels(
                    graph,
                    pos,
                    ax=ax,
                    font_size=8.5,
                    font_weight="bold",
                    labels={bus_name: _bus_label(bus_name) for bus_name in graph.nodes},
                )

                pair_counts: dict[tuple[str, str], int] = {}
                pair_index: dict[tuple[str, str], int] = {}
                for item in line_data:
                    pair = tuple(sorted((item["bus1"], item["bus2"])))
                    pair_counts[pair] = pair_counts.get(pair, 0) + 1

                for item in line_data:
                    pair = tuple(sorted((item["bus1"], item["bus2"])))
                    index = pair_index.get(pair, 0)
                    pair_index[pair] = index + 1
                    total = pair_counts[pair]
                    if total == 1:
                        rad = 0.0
                    else:
                        center = (total - 1) / 2.0
                        rad = (index - center) * 0.18

                    loading = item["loading"]
                    if loading is None or math.isnan(loading):
                        color = "#9e9e9e"
                        width = 1.2
                    else:
                        color = plt.cm.viridis(min(max(loading / 100.0, 0.0), 1.0))
                        width = 1.0 + min(max(loading, 0.0), 100.0) / 30.0

                    edge_patch = FancyArrowPatch(
                        posA=pos[item["bus1"]],
                        posB=pos[item["bus2"]],
                        arrowstyle="-",
                        connectionstyle=f"arc3,rad={rad}",
                        mutation_scale=10.0,
                        lw=width,
                        color=color,
                        alpha=0.9,
                        zorder=2,
                    )
                    ax.add_patch(edge_patch)

                    x1, y1 = pos[item["bus1"]]
                    x2, y2 = pos[item["bus2"]]

                load_counts: dict[str, int] = {}
                load_index: dict[str, int] = {}
                for item in load_data:
                    load_counts[item["bus"]] = load_counts.get(item["bus"], 0) + 1

                for item in load_data:
                    bus_name = item["bus"]
                    index = load_index.get(bus_name, 0)
                    load_index[bus_name] = index + 1
                    total = load_counts[bus_name]

                    bus_x, bus_y = pos[bus_name]
                    center_x, center_y = 0.0, 0.0
                    vec_x = bus_x - center_x
                    vec_y = bus_y - center_y
                    vec_len = math.hypot(vec_x, vec_y) or 1.0
                    unit_x = vec_x / vec_len
                    unit_y = vec_y / vec_len
                    perp_x = -unit_y
                    perp_y = unit_x

                    lateral = (index - (total - 1) / 2.0) * 0.10
                    arrow_start = (bus_x, bus_y)
                    arrow_end = (
                        bus_x + unit_x * 0.30 + perp_x * lateral,
                        bus_y + unit_y * 0.30 + perp_y * lateral,
                    )

                    load_arrow = FancyArrowPatch(
                        posA=arrow_start,
                        posB=arrow_end,
                        arrowstyle="-|>",
                        mutation_scale=12.0,
                        lw=1.2,
                        color="#d95f02",
                        alpha=0.95,
                        zorder=3,
                    )
                    ax.add_patch(load_arrow)

                    pass

                gen_counts: dict[str, int] = {}
                gen_index: dict[str, int] = {}
                for item in generator_data:
                    gen_counts[item["bus"]] = gen_counts.get(item["bus"], 0) + 1

                for item in generator_data:
                    bus_name = item["bus"]
                    index = gen_index.get(bus_name, 0)
                    gen_index[bus_name] = index + 1
                    total = gen_counts[bus_name]

                    bus_x, bus_y = pos[bus_name]
                    center_x, center_y = 0.0, 0.0
                    vec_x = bus_x - center_x
                    vec_y = bus_y - center_y
                    vec_len = math.hypot(vec_x, vec_y) or 1.0
                    unit_x = vec_x / vec_len
                    unit_y = vec_y / vec_len
                    perp_x = -unit_y
                    perp_y = unit_x

                    lateral = (index - (total - 1) / 2.0) * 0.10
                    arrow_start = (
                        bus_x + unit_x * 0.55 + perp_x * lateral,
                        bus_y + unit_y * 0.55 + perp_y * lateral,
                    )
                    arrow_end = (bus_x, bus_y)

                    gen_arrow = FancyArrowPatch(
                        posA=arrow_start,
                        posB=arrow_end,
                        arrowstyle="-|>",
                        mutation_scale=14.0,
                        lw=1.6,
                        color="#2ca25f",
                        alpha=0.95,
                        zorder=3,
                    )
                    ax.add_patch(gen_arrow)

                    pass

                summary_lines = [
                    f"Study case: {self.cfg.study_case}",
                    "",
                    f"Buses: {len(bus_names)}",
                    f"Branches: {len(line_data)}",
                    f"Loads: {len(load_data)}",
                    f"Generators: {len(generator_data)}",
                    "",
                    "Color scale: line loading",
                    "Edges are drawn without labels",
                    "to keep the topology readable.",
                    "",
                    "Generators",
                ]
                for item in generator_data:
                    p_txt = "n/a" if item["p_ini"] is None or math.isnan(item["p_ini"]) else f"{item['p_ini']:.2f}"
                    q_txt = "n/a" if item["q_ini"] is None or math.isnan(item["q_ini"]) else f"{item['q_ini']:.2f}"
                    summary_lines.append(f"{item['name']} -> {item['bus']}")
                    summary_lines.append(f"P={p_txt} MW, Q={q_txt} MVAr")
                summary_lines.append("")
                summary_lines.append("Loads")
                for item in load_data:
                    p_txt = "n/a" if item["p_ini"] is None or math.isnan(item["p_ini"]) else f"{item['p_ini']:.2f}"
                    q_txt = "n/a" if item["q_ini"] is None or math.isnan(item["q_ini"]) else f"{item['q_ini']:.2f}"
                    summary_lines.append(f"{item['name']} -> {item['bus']}")
                    summary_lines.append(f"P={p_txt} MW, Q={q_txt} MVAr")

                ax_info.text(
                    0.02,
                    0.98,
                    "\n".join(summary_lines),
                    transform=ax_info.transAxes,
                    va="top",
                    ha="left",
                    fontsize=8.2,
                    family="monospace",
                    bbox=dict(boxstyle="round,pad=0.5", facecolor="#fafafa", edgecolor="#d0d0d0"),
                )
            else:
                bus_list = sorted(bus_names)
                positions = {}
                radius = 1.0
                total_buses = len(bus_list)
                for index, bus_name in enumerate(bus_list):
                    angle = (2.0 * math.pi * index) / max(total_buses, 1)
                    positions[bus_name] = (radius * math.cos(angle), radius * math.sin(angle))

                fig, (ax, ax_info) = plt.subplots(
                    1,
                    2,
                    figsize=(22, 12),
                    gridspec_kw={"width_ratios": [4.6, 1.4]},
                )
                ax.set_title(f"PowerFactory Grid Topology — {self.cfg.study_case}")
                ax.axis("off")
                ax_info.axis("off")

                for bus_name, (x, y) in positions.items():
                    ax.scatter(x, y, s=900, color="#f7f2df", edgecolors="#333333", linewidths=0.9, zorder=3)
                    ax.text(x, y, _bus_label(bus_name), fontsize=7.5, fontweight="bold", ha="center", va="center", zorder=4)

                pair_counts: dict[tuple[str, str], int] = {}
                pair_index: dict[tuple[str, str], int] = {}
                for item in line_data:
                    pair = tuple(sorted((item["bus1"], item["bus2"])))
                    pair_counts[pair] = pair_counts.get(pair, 0) + 1

                for item in line_data:
                    pair = tuple(sorted((item["bus1"], item["bus2"])))
                    index = pair_index.get(pair, 0)
                    pair_index[pair] = index + 1
                    total = pair_counts[pair]
                    if total == 1:
                        rad = 0.0
                    else:
                        center = (total - 1) / 2.0
                        rad = (index - center) * 0.18

                    x1, y1 = positions[item["bus1"]]
                    x2, y2 = positions[item["bus2"]]
                    color = "#9e9e9e"
                    width = 1.2
                    loading = item["loading"]
                    if loading is not None and not math.isnan(loading):
                        color = plt.cm.viridis(min(max(loading / 100.0, 0.0), 1.0))
                        width = 1.0 + min(max(loading, 0.0), 100.0) / 30.0

                    ax.annotate(
                        "",
                        xy=(x2, y2),
                        xytext=(x1, y1),
                        arrowprops=dict(
                            arrowstyle="-",
                            color=color,
                            lw=width,
                            connectionstyle=f"arc3,rad={rad}",
                            alpha=0.9,
                        ),
                        zorder=2,
                    )

                    pass

                load_counts: dict[str, int] = {}
                load_index: dict[str, int] = {}
                for item in load_data:
                    load_counts[item["bus"]] = load_counts.get(item["bus"], 0) + 1

                for item in load_data:
                    bus_name = item["bus"]
                    index = load_index.get(bus_name, 0)
                    load_index[bus_name] = index + 1
                    total = load_counts[bus_name]

                    bus_x, bus_y = positions[bus_name]
                    center_x, center_y = 0.0, 0.0
                    vec_x = bus_x - center_x
                    vec_y = bus_y - center_y
                    vec_len = math.hypot(vec_x, vec_y) or 1.0
                    unit_x = vec_x / vec_len
                    unit_y = vec_y / vec_len
                    perp_x = -unit_y
                    perp_y = unit_x

                    lateral = (index - (total - 1) / 2.0) * 0.10
                    arrow_end = (
                        bus_x + unit_x * 0.30 + perp_x * lateral,
                        bus_y + unit_y * 0.30 + perp_y * lateral,
                    )

                    load_arrow = FancyArrowPatch(
                        posA=(bus_x, bus_y),
                        posB=arrow_end,
                        arrowstyle="-|>",
                        mutation_scale=12.0,
                        lw=1.2,
                        color="#d95f02",
                        alpha=0.95,
                        zorder=3,
                    )
                    ax.add_patch(load_arrow)

                    pass

                gen_counts: dict[str, int] = {}
                gen_index: dict[str, int] = {}
                for item in generator_data:
                    gen_counts[item["bus"]] = gen_counts.get(item["bus"], 0) + 1

                for item in generator_data:
                    bus_name = item["bus"]
                    index = gen_index.get(bus_name, 0)
                    gen_index[bus_name] = index + 1
                    total = gen_counts[bus_name]

                    bus_x, bus_y = positions[bus_name]
                    center_x, center_y = 0.0, 0.0
                    vec_x = bus_x - center_x
                    vec_y = bus_y - center_y
                    vec_len = math.hypot(vec_x, vec_y) or 1.0
                    unit_x = vec_x / vec_len
                    unit_y = vec_y / vec_len
                    perp_x = -unit_y
                    perp_y = unit_x

                    lateral = (index - (total - 1) / 2.0) * 0.10
                    arrow_start = (
                        bus_x + unit_x * 0.55 + perp_x * lateral,
                        bus_y + unit_y * 0.55 + perp_y * lateral,
                    )
                    arrow_end = (bus_x, bus_y)

                    gen_arrow = FancyArrowPatch(
                        posA=arrow_start,
                        posB=arrow_end,
                        arrowstyle="-|>",
                        mutation_scale=14.0,
                        lw=1.6,
                        color="#2ca25f",
                        alpha=0.95,
                        zorder=3,
                    )
                    ax.add_patch(gen_arrow)

                    pass

                summary_lines = [
                    f"Study case: {self.cfg.study_case}",
                    "",
                    f"Buses: {len(bus_names)}",
                    f"Branches: {len(line_data)}",
                    f"Loads: {len(load_data)}",
                    f"Generators: {len(generator_data)}",
                    "",
                    "Color scale: line loading",
                    "Edges are drawn without labels",
                    "to keep the topology readable.",
                    "",
                    "Generators",
                ]
                for item in generator_data:
                    p_txt = "n/a" if item["p_ini"] is None or math.isnan(item["p_ini"]) else f"{item['p_ini']:.2f}"
                    q_txt = "n/a" if item["q_ini"] is None or math.isnan(item["q_ini"]) else f"{item['q_ini']:.2f}"
                    summary_lines.append(f"{item['name']} -> {item['bus']}")
                    summary_lines.append(f"P={p_txt} MW, Q={q_txt} MVAr")
                summary_lines.append("")
                summary_lines.append("Loads")
                for item in load_data:
                    p_txt = "n/a" if item["p_ini"] is None or math.isnan(item["p_ini"]) else f"{item['p_ini']:.2f}"
                    q_txt = "n/a" if item["q_ini"] is None or math.isnan(item["q_ini"]) else f"{item['q_ini']:.2f}"
                    summary_lines.append(f"{item['name']} -> {item['bus']}")
                    summary_lines.append(f"P={p_txt} MW, Q={q_txt} MVAr")

                ax_info.text(
                    0.02,
                    0.98,
                    "\n".join(summary_lines),
                    transform=ax_info.transAxes,
                    va="top",
                    ha="left",
                    fontsize=8.2,
                    family="monospace",
                    bbox=dict(boxstyle="round,pad=0.5", facecolor="#fafafa", edgecolor="#d0d0d0"),
                )

            sm = plt.cm.ScalarMappable(cmap=plt.cm.viridis)
            sm.set_array([0.0, 100.0])
            cbar = fig.colorbar(sm, ax=ax, fraction=0.03, pad=0.02)
            cbar.set_label("Line loading [%]")

            fig.tight_layout()
            fig.savefig(graph_path, dpi=180, bbox_inches="tight")
            plt.close(fig)

            # Store structured data so downstream agents can use it directly
            # without interpreting the PNG image.
            self.grid_data = {
                "buses":      bus_characteristics,
                "lines":      line_data,
                "loads":      load_data,
                "generators": generator_data,
            }

            log.ok(
                f"Grid graph saved → {graph_path} "
                f"({len(line_data)} edges from {len(lines)} lines + {len(trf2_list)} trf2, "
                f"{len(load_data)} loads, {len(generator_data)} generators, {len(bus_names)} buses, {skipped} unresolved elements skipped)"
            )
            return True, graph_path

        except Exception as e:
            log.error(f"Grid graph export failed: {e}")
            return False, str(e)

    # ──────────────────────────────────────────────────────────────
    # STEP 4 — Run load flow
    # ──────────────────────────────────────────────────────────────

    def run_loadflow(self) -> tuple[bool, str]:
        log.section("STEP 3 — Load Flow (ComLdf)")
        try:
            ldf = self.app.GetFromStudyCase('ComLdf')
            err = ldf.Execute()
            if err:
                raise RuntimeError(f"ComLdf returned error code {err}")
            log.ok("Load flow converged successfully")
            return True, "Load flow OK"
        except Exception as e:
            log.error(f"Load flow failed: {e}")
            return False, str(e)

    # ──────────────────────────────────────────────────────────────
    # STEP 5 — Configure & run RMS simulation
    # ──────────────────────────────────────────────────────────────

    def run_rms_simulation(self) -> tuple[bool, str]:
        log.section("STEP 5 — RMS Simulation (ComInc + ComSim)")
        try:
            # -- Build fault events BEFORE initialisation -------------
            log.info(f"Applying fault at t={self.cfg.t_fault}s, clearing at t={self.cfg.t_clear}s")
            self._apply_fault_event()

            # -- Initialise simulation --------------------------------
            inc = self.app.GetFromStudyCase('ComInc')
            inc.iopt_sim   = 'rms'
            inc.iopt_show  = 0
            inc.iopt_adapt = 0
            inc.dtgrd      = self.cfg.dt_rms
            inc.start      = self.cfg.t_start
            self.app.EchoOff()
            err = inc.Execute()
            self.app.EchoOn()
            if err:
                raise RuntimeError(f"ComInc (initialisation) returned error code {err}")
            log.ok(f"Simulation initialised | dt={self.cfg.dt_rms}s")

            # -- Run simulation ---------------------------------------
            sim = self.app.GetFromStudyCase('ComSim')
            sim.tstop = self.cfg.t_end
            err = sim.Execute()
            if err:
                raise RuntimeError(f"ComSim returned error code {err}")
            log.ok(f"RMS simulation completed | t_end={self.cfg.t_end}s")
            return True, "RMS simulation OK"

        except Exception as e:
            log.error(f"RMS simulation failed: {e}")
            return False, str(e)

    def _apply_fault_event(self):
        """
        Clear all existing events, then create fault ON + clearance events.

        fault_type = "bus"  : EvtShc ON → EvtShc OFF (removes short-circuit)
        fault_type = "line" : EvtShc ON → EvtSwitch OPEN (trips the line)
        fault_type = "gen_switch" : EvtSwitch on selected generator
        """
        try:
            evt_folder = self.app.GetFromStudyCase('Simulation Events/Fault.IntEvt')
            if evt_folder is None:
                raise RuntimeError("Event folder not found: Simulation Events/Fault.IntEvt")

            # -- Clear existing events --------------------------------
            for obj in evt_folder.GetContents():
                obj.Delete()
            log.info("Existing simulation events cleared")

            raw_fault_type = str(getattr(self.cfg, "fault_type", "bus") or "bus")
            fault_type = raw_fault_type.strip().lower().replace("-", "_").replace(" ", "_")
            if fault_type in ("generator", "switch", "generator_switch"):
                fault_type = "gen_switch"

            if fault_type == "gen_switch":
                switch_element = (
                    getattr(self.cfg, "switch_element", "")
                    or getattr(self.cfg, "fault_element", "")
                )
                switch_time = float(
                    getattr(self.cfg, "t_switch", getattr(self.cfg, "switch_time", getattr(self.cfg, "t_fault", 1.0)))
                )

                raw_switch_state = getattr(self.cfg, "switch_state", getattr(self.cfg, "open_close", 0))
                if isinstance(raw_switch_state, str):
                    s = raw_switch_state.strip().lower()
                    if s in ("open", "trip", "off"):
                        switch_state = 0
                    elif s in ("close", "on"):
                        switch_state = 1
                    else:
                        switch_state = int(raw_switch_state)
                else:
                    switch_state = int(raw_switch_state)

                matches = self.app.GetCalcRelevantObjects(switch_element)
                if (not matches) and switch_element and ("." not in switch_element):
                    matches = self.app.GetCalcRelevantObjects(f"{switch_element}.ElmSym")
                if not matches:
                    all_gens = self.app.GetCalcRelevantObjects("*.ElmSym")
                    matches = [g for g in all_gens if getattr(g, "loc_name", "") == switch_element]
                if not matches:
                    raise RuntimeError(f"Switch target not found: {switch_element}")
                target = matches[0]

                # If a dedicated switch object exists for this generator name, prefer it.
                switch_obj_matches = self.app.GetCalcRelevantObjects(f"{target.loc_name}.StaSwitch")
                if switch_obj_matches:
                    target = switch_obj_matches[0]

                self.addSwitchEvent(target, switch_time, switch_state)
                action = "OPEN" if switch_state == 0 else "CLOSE"
                target_name = getattr(target, "loc_name", switch_element)
                log.info(f"EvtSwitch {action} → {target_name} at t={switch_time}s")
                return

            if fault_type not in ("bus", "line"):
                raise RuntimeError(f"Unsupported fault_type '{raw_fault_type}'. Use bus, line, or gen_switch.")

            # -- Faulted element --------------------------------------
            target = self.app.GetCalcRelevantObjects(self.cfg.fault_element)[0]

            # -- Short-circuit ON (same for both types) ---------------
            sc_on          = evt_folder.CreateObject('EvtShc', target.loc_name)
            sc_on.p_target = target
            sc_on.time     = self.cfg.t_fault
            sc_on.i_shc    = 0   # 3-phase fault
            log.info(f"EvtShc ON  → {self.cfg.fault_element} at t={self.cfg.t_fault}s")

            # -- Clearance (depends on fault_type) --------------------
            if fault_type == "line":
                # Trip the line: open its switch at t_clear
                self.addSwitchEvent(target, self.cfg.t_clear, 0)
                log.info(f"EvtSwitch OPEN → {self.cfg.fault_element} at t={self.cfg.t_clear}s")
            else:
                # Bus fault: remove short-circuit at t_clear
                sc_off          = evt_folder.CreateObject('EvtShc', target.loc_name)
                sc_off.p_target = target
                sc_off.time     = self.cfg.t_clear
                sc_off.i_shc    = 4   # clear fault
                log.info(f"EvtShc OFF → {self.cfg.fault_element} at t={self.cfg.t_clear}s")

        except Exception as e:
            log.warn(f"Could not create fault events automatically: {e}")
            log.warn("Continuing simulation without explicit fault — check your IntEvt folder")

    def addSwitchEvent(self, obj, sec, open_close):
        faultFolder = self.app.GetFromStudyCase("Simulation Events/Fault.IntEvt")
        if faultFolder is None:
            raise RuntimeError("Event folder not found: Simulation Events/Fault.IntEvt")
        event = faultFolder.CreateObject("EvtSwitch", obj.loc_name)
        if event is None:
            raise RuntimeError(f"Could not create EvtSwitch for target '{obj.loc_name}'")
        event.p_target = obj
        event.time = sec
        event.i_switch = open_close
        return event

    # ──────────────────────────────────────────────────────────────
    # STEP 6 — Export results to CSV
    # ──────────────────────────────────────────────────────────────

    def export_results_to_csv(self) -> tuple[bool, str]:
        log.section("STEP 6 — Export Results to CSV")
        try:
            os.makedirs(self.cfg.output_dir, exist_ok=True)

            filename = os.path.join(
                self.cfg.output_dir,
                f"{self.cfg.run_label}_RMS.csv"
            )

            # -- Use ComRes (PowerFactory built-in CSV exporter) ------
            comRes = self.app.GetFromStudyCase("ComRes")
            comRes.pResult  = self.app.GetFromStudyCase(self.cfg.result_name)
            comRes.f_name   = filename
            comRes.iopt_sep = 0   # use custom separators below
            comRes.col_Sep  = ";" # column separator
            comRes.dec_Sep  = "." # decimal separator
            comRes.iopt_exp = 6   # export format: CSV with time column
            comRes.iopt_csel = 0  # all columns
            comRes.iopt_vars = 0  # all variables
            comRes.iopt_tsel = 0  # full time range
            comRes.iopt_rscl = 0  # no rescaling
            err = comRes.Execute()
            if err:
                raise RuntimeError(f"ComRes.Execute() returned error code {err}")

            log.ok(f"CSV saved → {filename}")
            return True, filename

        except Exception as e:
            log.error(f"CSV export failed: {e}")
            return False, str(e)

    # ──────────────────────────────────────────────────────────────
    # PIPELINE — run all steps in sequence
    # ──────────────────────────────────────────────────────────────

    def run_pipeline(self) -> dict:
        """
        Execute the full pipeline and return a status report dict.
        Each step is guarded: a failure stops the pipeline early.
        """
        report = {
            "connect":          None,
            "activate_case":    None,
            "load_flow":        None,
            "network_graph":    None,
            "rms_simulation":   None,
            "csv_export":       None,
            "csv_path":         None,
            "success":          False,
        }

        steps = [
            ("connect",        self.connect),
            ("activate_case",  self.activate_study_case),
            ("load_flow",      self.run_loadflow),
            ("network_graph",  self.export_grid_graph),
            ("rms_simulation", self.run_rms_simulation),
            ("csv_export",     self.export_results_to_csv),
        ]

        for key, fn in steps:
            ok, msg = fn()
            report[key] = {"ok": ok, "msg": msg}
            if not ok:
                log.error(f"Pipeline stopped at step '{key}': {msg}")
                return report

        report["csv_path"]  = report["csv_export"]["msg"]
        report["grid_data"] = self.grid_data
        report["success"]   = True
        log.section("PIPELINE COMPLETE")
        log.ok(f"All steps passed. Results → {report['csv_path']}")
        return report


# ══════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":

    # ── Load config from JSON (edit simulation_config.json, not this file)
    _cfg_path = os.path.join(os.path.dirname(__file__), "simulation_config.json")
    cfg = SimulationConfig.from_json(_cfg_path)

    agent  = DIgSILENTAgent(cfg)
    report = agent.run_pipeline()

    # ── Print final summary ────────────────────────────────────────
    print("\n" + "═" * 60)
    print("  PIPELINE REPORT")
    print("═" * 60)
    for step, result in report.items():
        if isinstance(result, dict):
            status = "✅" if result["ok"] else "❌"
            print(f"  {status}  {step:<20} {result['msg']}")
    print(f"\n  Overall success: {'✅ YES' if report['success'] else '❌ NO'}")
    if report["csv_path"]:
        print(f"  CSV output:      {report['csv_path']}")
    print("═" * 60)

