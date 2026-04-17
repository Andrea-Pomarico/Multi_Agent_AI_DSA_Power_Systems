"""
Multi-Agent Dynamic Security Assessment Pipeline
============================================
Orchestrates a chain of specialized agents to simulate, analyse, and report
on dynamic simulation power-system transient stability using DIgSILENT
PowerFactory as the simulation engine and an LLM backend for narrative
reporting.

Pipeline — single-case
----------------------
  Agent 0  intake_agent          parse user input or load config cases
  Agent 1  simulation_agent      run PowerFactory RMS simulation → CSV
                                  (also exports pre-fault grid graph: V, angle,
                                   Ikss, line loadings, P/Q flows)
  Agent 2  analysis_agent        compute numerical KPIs from CSV
  Agent 3  plot_agent            generate time-series PNG plots
  Agent 4  summary_agent         LLM narrative summary of dynamics
  Agent 5  review_agent          LLM cross-check of summary vs KPIs  (optional)
  Agent 6  final_report_agent    LLM polished final engineering report (optional)
  Agent 7  comparison_agent      LLM cross-case comparative report   (multi-case)
  Agent 8  presentation_agent    build PowerPoint from report + plots
  Agent 9  mitigation_agent      LLM grid-topology analysis + ranked
                                  mitigation measures (optional)
Usage
-----
  python Multi_Agent_RMS_google.py

Configuration
-------------
  Edit simulation_config.json to set PowerFactory paths, fault parameters,
  LLM provider / API keys, and output directories.
  See simulation_config.example.json for a template with placeholder keys.
"""

import os
import json
import time
from datetime import datetime

from Agent_DIgSILENT import SimulationConfig
from llm_client import MODEL_FAST, configure_llm_from_config, get_llm_runtime_info
from prompt_loader import load_prompts

from agent_intake import intake_agent
from agent_simulation import simulation_agent
from agent_analysis import analysis_agent
from agent_plot import (
    plot_agent,
    plot_voltage_case_comparison,
    plot_speed_case_comparison,
)
from agent_llm_reporting import (
    summary_agent,
    review_agent,
    final_report_agent,
    comparison_agent,
)
from agent_presentation import presentation_agent
from agent_mitigation import mitigation_agent
from report_utils import (
    save_pipeline_log,
    save_summary_csv,
    save_kpi_csv,
    save_report_docx,
)


# ══════════════════════════════════════════════════════════════════
# SINGLE-CASE PIPELINE ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════

def run_rms_pipeline(cfg: SimulationConfig,
                     emit_final_report: bool = True,
                     emit_final_presentation: bool = True) -> dict:
    """
    Execute the full single-case RMS stability analysis pipeline.

    Each step is timed and appended to an in-memory log that is flushed to
    CSV at the end.  On simulation failure the pipeline aborts early and
    returns whatever the simulation agent reported.

    Parameters
    ----------
    cfg : SimulationConfig
        All simulation parameters (PowerFactory paths, fault timing, output
        directory, LLM flags, etc.).

    Returns
    -------
    dict
        Keys: success, csv_path, run_dir, numerics, parsed, statistics,
              plots, log_csv, kpi_csv, summary_csv, summary, improvements,
              final_report, combined_report, final_report_docx,
              presentation_pptx.
    """
    print("\n" + "═" * 60)
    print("  MULTI-AGENT RMS PIPELINE")
    print("═" * 60)
    t0 = time.time()

    # Each run gets its own timestamped sub-folder so outputs never collide.
    run_ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = os.path.join(cfg.output_dir, f"{cfg.run_label}_{run_ts}")
    os.makedirs(run_dir, exist_ok=True)
    cfg.output_dir = run_dir
    print(f"[INFO] Run folder → {run_dir}")

    log_rows: list[dict] = []

    def _log(step: str, ok: bool, msg: str, t_start: float) -> None:
        log_rows.append({
            "timestamp":  datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "step":       step,
            "status":     "OK" if ok else "ERROR",
            "message":    msg,
            "duration_s": round(time.time() - t_start, 3),
        })

    # ── Agent 1: PowerFactory RMS simulation ─────────────────────
    t1     = time.time()
    report = simulation_agent(cfg)
    sim_ok = bool(report.get("success")) if isinstance(report, dict) else False

    # Extract the grid graph PNG path and structured grid data produced
    # during load-flow (both stored by DIgSILENTAgent in the report).
    _ng        = report.get("network_graph") if isinstance(report, dict) else None
    graph_path = (_ng.get("msg", "") if isinstance(_ng, dict) else "") or ""
    if graph_path and not os.path.isfile(graph_path):
        graph_path = ""
    grid_data  = report.get("grid_data", {}) if isinstance(report, dict) else {}

    csv_export = report.get("csv_export") if isinstance(report, dict) else None
    if isinstance(csv_export, dict):
        sim_msg = csv_export.get("msg", "")
    elif isinstance(report, dict):
        sim_msg = report.get("error", "") or "Simulation failed before CSV export"
    else:
        sim_msg = "Simulation agent returned invalid report"
    _log("simulation", sim_ok, sim_msg, t1)

    if not sim_ok:
        if not isinstance(report, dict):
            report = {
                "success":  False,
                "error":    "Simulation agent returned invalid report",
                "csv_path": None,
            }
        _log("pipeline", False, "Aborted — simulation failed", t0)
        save_pipeline_log(log_rows, run_dir, cfg.run_label)
        print("[PIPELINE ABORTED] Simulation failed — no CSV to analyse.")
        return report

    # ── Agent 2: Numerical KPI analysis ──────────────────────────
    t2             = time.time()
    numerics, stats_txt = analysis_agent(
        report["csv_path"],
        cfg.t_fault,
        cfg.t_clear,
        fault_type=getattr(cfg, "fault_type", "bus"),
        switch_element=getattr(cfg, "switch_element", ""),
        fault_element=getattr(cfg, "fault_element", ""),
        switch_state=getattr(cfg, "switch_state", 0),
        t_switch=getattr(cfg, "t_switch", cfg.t_fault),
    )
    _log("analysis", bool(numerics),
         f"{len(numerics.get('voltages', {}))} buses  "
         f"{len(numerics.get('speeds', {}))} gens", t2)

    # ── Agent 3: Time-series plots ────────────────────────────────
    t3         = time.time()
    plot_paths = plot_agent(numerics, cfg)
    _log("plots", bool(plot_paths), f"{len(plot_paths)} PNG files saved", t3)

    # ── Agent 4: LLM narrative summary ───────────────────────────
    t4      = time.time()
    summary = summary_agent(numerics, cfg.study_case)
    _log("llm_summary", True, f"model={MODEL_FAST}", t4)

    # ── Agent 5: LLM review — controlled by run_review_agent in config ──
    t5 = time.time()
    if int(getattr(cfg, "run_review_agent", 0)) == 1:
        improvements = review_agent(numerics, summary)
        _log("review", True, f"model={MODEL_FAST}", t5)
    else:
        improvements = "Review step skipped."
        _log("review", True, "Skipped (run_review_agent=0)", t5)

    # ── Agent 6: LLM final report — controlled by run_final_report_agent ──
    t6 = time.time()
    if int(getattr(cfg, "run_final_report_agent", 0)) == 1:
        final_report = final_report_agent(summary, improvements, cfg.study_case)
        _log("final_report", True, f"model={MODEL_FAST}", t6)
    else:
        final_report = summary
        _log("final_report", True, "Skipped (run_final_report_agent=0)", t6)

    # ── Agent 9: Mitigation — controlled by run_mitigation_agent ────
    t9              = time.time()
    mitigation_txt  = ""
    mitigation_path = ""
    if int(getattr(cfg, "run_mitigation_agent", 0)) == 1:
        mitigation_txt = mitigation_agent(
            final_report=final_report,
            numerics=numerics,
            grid_data=grid_data,
            study_case=cfg.study_case,
        )
        mitigation_path = os.path.join(run_dir, f"{cfg.run_label}_mitigation.txt")
        with open(mitigation_path, "w", encoding="utf-8") as _mf:
            _mf.write(mitigation_txt)
        _log("mitigation", True, f"Saved → {mitigation_path}", t9)
        print(f"[OK] Mitigation    → {mitigation_path}")
    else:
        _log("mitigation", True, "Skipped (run_mitigation_agent=0)", t9)

    # ── Build combined report (stability analysis + mitigation) ──
    # When Agent 9 ran, its output is appended so final artefacts
    # contain both dynamics analysis and mitigation.
    if mitigation_txt:
        combined_report = (
            final_report
            + "\n\n"
            + "═" * 60 + "\n"
            + "  MITIGATION ANALYSIS\n"
            + "═" * 60 + "\n"
            + mitigation_txt
        )
    else:
        combined_report = final_report

    # ── Flush all outputs to disk ─────────────────────────────────
    _log("pipeline", True, f"Complete in {time.time() - t0:.1f}s", t0)
    log_path          = save_pipeline_log(log_rows, run_dir, cfg.run_label)
    kpi_path          = save_kpi_csv(numerics, run_dir, cfg.run_label)
    summary_path      = save_summary_csv(summary, cfg, run_dir, cfg.run_label)
    final_report_path = ""
    if emit_final_report:
        final_report_path = save_report_docx(
            combined_report, run_dir, cfg.run_label, plot_paths=plot_paths
        )
        print(f"[OK] Final report (Word) → {final_report_path}")
    else:
        print("[INFO] Final report export skipped for this case (aggregated run output mode)")

    print(f"[OK] Pipeline log  → {log_path}")
    print(f"[OK] KPI summary   → {kpi_path}")
    print(f"[OK] LLM summary   → {summary_path}")

    # ── Agent 8: PowerPoint presentation ─────────────────────────
    ppt_ok = False
    pptx_path = ""
    if emit_final_presentation:
        ppt_ok, pptx_path, ppt_msg = presentation_agent(
            report_text=combined_report,
            study_case=cfg.study_case,
            out_dir=run_dir,
            label=cfg.run_label,
            plot_paths=plot_paths,
        )
        if ppt_ok:
            print(f"[OK] Presentation  → {pptx_path}")
        else:
            print(f"[WARN] Presentation skipped: {ppt_msg}")
    else:
        print("[INFO] Presentation export skipped for this case (aggregated run output mode)")

    elapsed = time.time() - t0
    print(f"\n{'═' * 60}")
    print(f"  PIPELINE COMPLETE  ({elapsed:.1f}s)")
    print(f"  Run folder : {run_dir}")
    print(f"  Plots      : {len(plot_paths)}")
    print("═" * 60)

    report.update({
        "run_dir":           run_dir,
        "numerics":          {k: v for k, v in numerics.items() if k != "_parsed"},
        "parsed":            numerics.get("_parsed", {}),
        "statistics":        stats_txt,
        "plots":             plot_paths,
        "log_csv":           log_path,
        "kpi_csv":           kpi_path,
        "summary_csv":       summary_path,
        "summary":           summary,
        "improvements":      improvements,
        "final_report":      final_report,
        "combined_report":   combined_report,
        "final_report_docx": final_report_path,
        "presentation_pptx": pptx_path if ppt_ok else "",
        "mitigation_report": mitigation_txt,
        "mitigation_txt":    mitigation_path,
        "graph_path":        graph_path,
    })
    return report


# ══════════════════════════════════════════════════════════════════
# MULTI-CASE PIPELINE ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════

def run_rms_multi_case_pipeline(cases_list: list[dict],
                                base_cfg: SimulationConfig) -> dict:
    """
    Run multiple fault-scenario case studies in sequence and compare results.

    Each case goes through the full single-case pipeline.  After all cases
    finish, cross-case comparison plots and an LLM comparative report are
    generated and saved alongside the individual case folders.

    Parameters
    ----------
    cases_list : list[dict]
        Each element is a case-spec dict (keys: case_name, fault_type,
        fault_element, t_fault, t_clear, …).  See simulation_config.json.
    base_cfg : SimulationConfig
        Template config — case-specific fields override it per case.

    Returns
    -------
    dict
          Keys: case_results, metadata, comparison_report,
              comparison_report_path, final_report_path,
              final_presentation_path.
    """
    print("\n" + "═" * 60)
    print("  MULTI-AGENT RMS PIPELINE (MULTI-CASE)")
    print("═" * 60)
    t0 = time.time()

    run_ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
    multi_dir = os.path.join(base_cfg.output_dir, f"multi_case_comparison_{run_ts}")
    os.makedirs(multi_dir, exist_ok=True)
    print(f"[INFO] Multi-case output folder → {multi_dir}")

    all_results: dict = {
        "case_results": {},
        "metadata": {
            "timestamp":  run_ts,
            "num_cases":  len(cases_list),
            "start_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        },
    }

    for i, case_spec in enumerate(cases_list, start=1):
        case_name = case_spec.get("case_name", f"Case_{i}")
        print(f"\n[{'─' * 50}]")
        print(f"[CASE {i}/{len(cases_list)}]: {case_name}")
        print(f"[{'─' * 50}]")

        cfg = _build_case_config(base_cfg, case_spec, multi_dir)
        os.makedirs(cfg.output_dir, exist_ok=True)

        case_report = run_rms_pipeline(
            cfg,
            emit_final_report=False,
            emit_final_presentation=False,
        )
        all_results["case_results"][case_name] = {
            "config": {
                "fault_type":    cfg.fault_type,
                "fault_element": cfg.fault_element,
                "switch_element": cfg.switch_element,
                "t_switch":      cfg.t_switch,
                "switch_state":  cfg.switch_state,
                "t_fault":       cfg.t_fault,
                "t_clear":       cfg.t_clear,
            },
            "numerics":          case_report.get("numerics", {}),
            "parsed":            case_report.get("parsed", {}),
            "summary":           case_report.get("summary", ""),
            "final_report":      case_report.get("final_report", ""),
            "combined_report":   case_report.get("combined_report", ""),
            "mitigation_report": case_report.get("mitigation_report", ""),
            "plots":             case_report.get("plots", []),
            "run_dir":           case_report.get("run_dir", ""),
        }

    case_names = list(all_results["case_results"].keys())
    comparison_report = ""
    comparison_plots: list[str] = []
    if len(case_names) >= 2:
        comparison_report, comparison_plots = _run_comparison(
            all_results, case_names, multi_dir, base_cfg
        )

    _save_multi_case_final_outputs(
        all_results=all_results,
        case_names=case_names,
        multi_dir=multi_dir,
        comparison_report=comparison_report,
        comparison_plots=comparison_plots,
    )

    elapsed = time.time() - t0
    all_results["metadata"]["end_time"]         = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    all_results["metadata"]["total_duration_s"] = elapsed

    print(f"\n{'═' * 60}")
    print(f"  MULTI-CASE PIPELINE COMPLETE")
    print(f"  Cases run    : {len(case_names)}")
    print(f"  Output folder: {multi_dir}")
    print(f"  Total time   : {elapsed:.1f}s")
    print("═" * 60)
    return all_results


# ══════════════════════════════════════════════════════════════════
# INTERNAL HELPERS
# ══════════════════════════════════════════════════════════════════

def _build_case_config(base_cfg: SimulationConfig, case_spec: dict,
                       parent_dir: str) -> SimulationConfig:
    """
    Clone *base_cfg* and apply case-specific overrides from *case_spec*.

    Normalises fault_type aliases (e.g. "generator" → "gen_switch") and
    resolves switch_state string values ("open" / "close").
    """
    cfg = SimulationConfig()

    # Copy all shared base fields so each case inherits common setup.
    cfg.project_path    = base_cfg.project_path
    cfg.base_study_case = base_cfg.base_study_case
    cfg.result_name     = base_cfg.result_name
    cfg.signals         = base_cfg.signals
    cfg.t_start         = base_cfg.t_start
    cfg.t_end           = base_cfg.t_end
    cfg.dt_rms          = base_cfg.dt_rms
    cfg.run_review_agent = int(getattr(base_cfg, "run_review_agent", 0))
    cfg.run_final_report_agent = int(getattr(base_cfg, "run_final_report_agent", 0))
    cfg.run_mitigation_agent = int(getattr(base_cfg, "run_mitigation_agent", 0))
    cfg.switch_element  = getattr(base_cfg, "switch_element", "")
    cfg.t_switch        = float(getattr(base_cfg, "t_switch", base_cfg.t_fault))
    cfg.switch_state    = int(getattr(base_cfg, "switch_state", 0))

    case_name      = case_spec.get("case_name", "Case")
    cfg.study_case = case_name
    cfg.run_label  = case_name
    cfg.output_dir = os.path.join(parent_dir, case_name)

    # Normalise fault_type: accept several aliases and map to canonical values.
    fault_type_raw = str(case_spec.get("fault_type", "bus")).strip().lower()
    fault_type_raw = fault_type_raw.replace("-", "_").replace(" ", "_")
    if fault_type_raw in ("generator", "switch", "generator_switch"):
        fault_type_raw = "gen_switch"
    cfg.fault_type = fault_type_raw

    cfg.fault_element  = case_spec.get("fault_element", base_cfg.fault_element)
    cfg.switch_element = case_spec.get(
        "switch_element", case_spec.get("switch_target", cfg.switch_element)
    )
    cfg.t_switch = float(
        case_spec.get("t_switch", case_spec.get("switch_time", cfg.t_switch))
    )

    raw_state = case_spec.get("switch_state", case_spec.get("open_close", cfg.switch_state))
    if isinstance(raw_state, str):
        s = raw_state.strip().lower()
        if s in ("open", "trip", "off"):
            cfg.switch_state = 0
        elif s in ("close", "on"):
            cfg.switch_state = 1
        else:
            cfg.switch_state = int(raw_state)
    else:
        cfg.switch_state = int(raw_state)

    cfg.t_start = float(case_spec.get("t_start", base_cfg.t_start))
    cfg.t_fault = float(case_spec.get("t_fault", base_cfg.t_fault))
    cfg.t_clear = float(case_spec.get("t_clear", base_cfg.t_clear))
    cfg.t_end   = float(case_spec.get("t_end",   base_cfg.t_end))
    cfg.dt_rms  = float(case_spec.get("dt_rms",  base_cfg.dt_rms))

    # Generator-switch faults allow fault_element / t_fault as target / time aliases.
    if cfg.fault_type == "gen_switch":
        if not cfg.switch_element:
            cfg.switch_element = cfg.fault_element
        if "t_switch" not in case_spec:
            cfg.t_switch = cfg.t_fault

    return cfg


def _run_comparison(all_results: dict, case_names: list[str],
                    multi_dir: str, base_cfg: SimulationConfig) -> tuple[str, list[str]]:
    """
    Generate cross-case comparison plots and an LLM comparative report.

    Called after all individual cases complete.  Mutates *all_results* in
    place, adding comparison_report and comparison_report_path.
    """
    results_dict = {
        name: all_results["case_results"][name]["numerics"]
        for name in case_names
    }

    voltage_plots = plot_voltage_case_comparison(
        all_results["case_results"], multi_dir, label="multi_case"
    )
    speed_plots = plot_speed_case_comparison(
        all_results["case_results"], multi_dir, label="multi_case"
    )
    comparison_plots = voltage_plots + speed_plots

    if voltage_plots:
        all_results["voltage_comparison_plots"] = voltage_plots
    if speed_plots:
        all_results["speed_comparison_plots"] = speed_plots

    comp_report = comparison_agent(results_dict, case_names)
    all_results["comparison_report"] = comp_report

    comp_txt_path = os.path.join(multi_dir, "comparison_report.txt")
    with open(comp_txt_path, "w", encoding="utf-8") as f:
        f.write("═" * 60 + "\n")
        f.write("  MULTI-CASE COMPARISON REPORT\n")
        f.write("═" * 60 + "\n\n")
        f.write(comp_report + "\n")
    all_results["comparison_report_path"] = comp_txt_path
    print(f"[OK] Comparison report → {comp_txt_path}")

    return comp_report, comparison_plots


def _save_multi_case_final_outputs(all_results: dict,
                                   case_names: list[str],
                                   multi_dir: str,
                                   comparison_report: str,
                                   comparison_plots: list[str]) -> None:
    """
    Save a single final report + single final presentation for the entire
    multi-case run.
    """
    report_parts: list[str] = [
        "═" * 60,
        "  FINAL RMS DYNAMIC STABILITY REPORT",
        "═" * 60,
        "",
        "1. PRESENTATION OF THE SCENARIO",
        "",
    ]

    for idx, case_name in enumerate(case_names, start=1):
        case_data = all_results["case_results"].get(case_name, {}) or {}
        cfg = case_data.get("config", {}) or {}
        report_parts.extend([
            f"Case {idx}: {case_name}",
            f"  fault_type={cfg.get('fault_type', 'n/a')}",
            f"  fault_element={cfg.get('fault_element', 'n/a')}",
            f"  switch_element={cfg.get('switch_element', 'n/a')}",
            f"  switch_state={cfg.get('switch_state', 'n/a')}",
            f"  t_fault={cfg.get('t_fault', 'n/a')} s",
            f"  t_clear={cfg.get('t_clear', 'n/a')} s",
            "",
        ])

    report_parts.extend([
        "2. VOLTAGE VIOLATIONS AND PLOTS",
        "",
    ])

    for idx, case_name in enumerate(case_names, start=1):
        case_data = all_results["case_results"].get(case_name, {}) or {}
        numerics = case_data.get("numerics", {}) or {}
        voltages = numerics.get("voltages", {}) or {}
        speeds = numerics.get("speeds", {}) or {}
        case_plot_paths = case_data.get("plots", []) or []

        report_parts.append(f"Case {idx}: {case_name}")
        violations = []
        for bus, kpis in sorted(voltages.items()):
            if bool(kpis.get("out_of_limit_after_check")):
                violations.append(
                    "  "
                    + f"{bus}: post_min={kpis.get('post_min', float('nan')):.4f} p.u. "
                    + f"(t={kpis.get('t_post_min', float('nan')):.3f}s), "
                    + f"post_max={kpis.get('post_max', float('nan')):.4f} p.u. "
                    + f"(t={kpis.get('t_post_max', float('nan')):.3f}s), "
                    + f"final={kpis.get('final_val', float('nan')):.4f} p.u."
                )

        if violations:
            report_parts.append("  Voltage violations after clearing (0.9-1.1 p.u. band):")
            report_parts.extend(violations)
        else:
            report_parts.append("  No post-clear voltage violations (all buses within 0.9-1.1 p.u.).")

        voltage_plot_names = []
        for p in case_plot_paths:
            if isinstance(p, str) and os.path.isfile(p):
                stem = os.path.splitext(os.path.basename(p))[0].lower()
                if "voltage" in stem:
                    voltage_plot_names.append(os.path.basename(p))
        if voltage_plot_names:
            report_parts.append("  Voltage plots:")
            for name in sorted(dict.fromkeys(voltage_plot_names)):
                report_parts.append(f"    - {name}")
        else:
            report_parts.append("  Voltage plots: none detected for this case.")

        speed_plot_names = []
        for p in case_plot_paths:
            if isinstance(p, str) and os.path.isfile(p):
                stem = os.path.splitext(os.path.basename(p))[0].lower()
                if "speed" in stem:
                    speed_plot_names.append(os.path.basename(p))
        if speed_plot_names:
            report_parts.append("  Speed plots:")
            for name in sorted(dict.fromkeys(speed_plot_names)):
                report_parts.append(f"    - {name}")
        else:
            if speeds:
                report_parts.append("  Speed plots: available in the run outputs.")
            else:
                report_parts.append("  Speed plots: no generator speed signals available.")
        report_parts.append("")

    if comparison_report:
        report_parts.extend([
            "Multi-case comparison note:",
            comparison_report,
            "",
        ])

    report_parts.extend([
        "3. PROPOSED MITIGATION",
        "",
    ])

    for idx, case_name in enumerate(case_names, start=1):
        case_data = all_results["case_results"].get(case_name, {}) or {}
        mitigation_text = case_data.get("mitigation_report", "") or ""
        report_parts.append(f"Case {idx}: {case_name}")
        if mitigation_text.strip():
            report_parts.append(mitigation_text)
        else:
            report_parts.append("Mitigation agent output not available (agent may be disabled).")
        report_parts.append("")

    final_text = "\n".join(report_parts).strip() + "\n"

    all_plot_paths: list[str] = []
    for case_name in case_names:
        case_plots = (all_results["case_results"].get(case_name, {}) or {}).get("plots", [])
        for p in case_plots:
            if isinstance(p, str) and os.path.isfile(p):
                stem = os.path.splitext(os.path.basename(p))[0].lower()
                if "voltage" in stem or "speed" in stem:
                    all_plot_paths.append(p)
    for p in comparison_plots:
        if isinstance(p, str) and os.path.isfile(p):
            stem = os.path.splitext(os.path.basename(p))[0].lower()
            if "voltage" in stem or "speed" in stem:
                all_plot_paths.append(p)

    dedup_plots = list(dict.fromkeys(all_plot_paths))

    final_report_path = save_report_docx(
        final_report=final_text,
        out_dir=multi_dir,
        label="final_multi_case",
        plot_paths=dedup_plots,
    )
    all_results["final_report_path"] = final_report_path
    print(f"[OK] Final consolidated report (Word) → {final_report_path}")
    ppt_ok, pptx_path, ppt_msg = presentation_agent(
        report_text=final_text,
        study_case="Final run summary",
        out_dir=multi_dir,
        label="final_multi_case",
        plot_paths=dedup_plots,
    )
    if ppt_ok:
        all_results["final_presentation_path"] = pptx_path
        print(f"[OK] Final consolidated presentation → {pptx_path}")
    else:
        print(f"[WARN] Final consolidated presentation skipped: {ppt_msg}")


def _cases_from_config_json(cfg_path: str) -> list[dict]:
    """
    Read the 'cases' (or legacy 'faults') list from simulation_config.json.

    Returns an empty list when the key is absent or the file cannot be parsed,
    allowing the caller to fall back to the interactive intake agent.
    """
    try:
        with open(cfg_path, "r", encoding="utf-8") as f:
            raw = json.load(f)
    except Exception:
        return []

    if not isinstance(raw, dict):
        return []

    cases = raw.get("cases") or raw.get("faults")
    if not isinstance(cases, list):
        return []

    out: list[dict] = []
    for i, case in enumerate(cases, start=1):
        if not isinstance(case, dict):
            continue

        fault_type = str(case.get("fault_type", "")).strip().lower()
        fault_type = fault_type.replace("-", "_").replace(" ", "_")
        if fault_type in ("generator", "switch", "generator_switch"):
            fault_type = "gen_switch"
        if not fault_type:
            continue

        if fault_type == "gen_switch":
            has_target = bool(
                case.get("switch_element") or case.get("switch_target")
                or case.get("fault_element")
            )
            has_time = (
                "t_switch" in case or "switch_time" in case or "t_fault" in case
            )
            if not (has_target and has_time):
                continue
        elif "fault_element" not in case:
            continue

        normalized = dict(case)
        normalized.setdefault("case_name", f"Case_{i}")
        out.append(normalized)

    return out


# ══════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    # Load agent system-prompts from CSV before any LLM call.
    load_prompts()

    # Resolve config path relative to this script so the pipeline can be
    # launched from any working directory.
    _cfg_path = os.path.join(os.path.dirname(__file__), "simulation_config.json")

    with open(_cfg_path, "r", encoding="utf-8") as _f:
        _raw_cfg = json.load(_f)

    configure_llm_from_config(_raw_cfg)
    _info = get_llm_runtime_info()
    print(
        f"[INFO] LLM provider : {_info['provider']}\n"
        f"[INFO] Model (fast) : {_info['model_fast']}\n"
        f"[INFO] Model (smart): {_info['model_smart']}"
    )

    cfg       = SimulationConfig.from_json(_cfg_path)
    cfg_cases = _cases_from_config_json(_cfg_path)

    if cfg_cases:
        # Cases are pre-defined in JSON — skip interactive intake.
        print(f"\n[INFO] Loaded {len(cfg_cases)} case(s) from simulation_config.json")
        result_cfg, cases_list = cfg, cfg_cases
    else:
        # No pre-defined cases — ask the user interactively.
        result_cfg, cases_list = intake_agent(cfg)

    if cases_list:
        print(f"\n[INFO] Running in MULTI-CASE mode ({len(cases_list)} cases)")
        run_rms_multi_case_pipeline(cases_list, result_cfg or cfg)
    else:
        print("\n[INFO] Running in SINGLE-CASE mode (consolidated output)")
        single_cfg = result_cfg or cfg
        single_case = {
            "case_name": single_cfg.run_label or "Case_1",
            "fault_type": single_cfg.fault_type,
            "fault_element": single_cfg.fault_element,
            "switch_element": getattr(single_cfg, "switch_element", ""),
            "t_switch": getattr(single_cfg, "t_switch", single_cfg.t_fault),
            "switch_state": getattr(single_cfg, "switch_state", 0),
            "t_fault": single_cfg.t_fault,
            "t_clear": single_cfg.t_clear,
            "t_end": single_cfg.t_end,
            "dt_rms": single_cfg.dt_rms,
        }
        run_rms_multi_case_pipeline([single_case], single_cfg)
