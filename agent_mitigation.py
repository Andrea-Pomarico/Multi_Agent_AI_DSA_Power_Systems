"""
Agent 9 — Mitigation Agent
===========================
Two-stage LLM agent that uses the structured grid data (the same data used
to draw the topology graph) together with the stability report to propose
concrete engineering countermeasures.

Why structured data instead of the PNG image
--------------------------------------------
Sending the PNG to a vision model is fragile:
  • Requires a vision-capable model (not available on all providers/tiers).
  • Interpreting numbers from a rendered image introduces errors.
  • The structured data is already available in memory with full precision.

The grid_data dict (produced by DIgSILENTAgent.export_grid_graph) contains:
  buses      — {bus_key: {name, u [p.u.], phiu [deg], ikss [kA]}}
  lines      — [{name, bus1, bus2, loading [%], p_flow [MW]}]
  loads      — [{name, bus, p_ini [MW], q_ini [MVAr]}]
  generators — [{name, bus, p_ini [MW], q_ini [MVAr]}]

Stage A — Grid Analysis
    Builds a rich, structured text block from grid_data and asks the LLM to
    identify structural vulnerabilities (voltage deviations, overloaded
    branches, weak fault-current buses, network bottlenecks).

Stage B — Mitigation Plan
    Merges Stage A findings + final stability report + KPIs into a ranked
    mitigation report:
      1. IMMEDIATE ACTIONS
      2. SHORT-TERM MEASURES
      3. LONG-TERM RECOMMENDATIONS

Public API
----------
  mitigation_agent(final_report, numerics, grid_data, study_case) → str
"""

import math

import numpy as np

from llm_client import MODEL_SMART, run_agent


# ══════════════════════════════════════════════════════════════════
# SYSTEM PROMPTS
# ══════════════════════════════════════════════════════════════════

_GRAPH_ANALYSIS_PROMPT = """\
You are a senior power-systems network analyst at a Transmission System Operator.
You are given structured pre-fault load-flow and short-circuit data for every
bus, branch, load, and generator in the network.

Your task — analyse the data and identify:

1. VOLTAGE PROFILE
   List every bus with V < 0.95 p.u. or V > 1.05 p.u.
   Note buses close to the statutory limits (0.95–1.05 p.u.) as at-risk.

2. VOLTAGE ANGLE SPREAD
   Flag buses with large voltage angles (> ±20°) — they indicate heavy
   power transfer that stresses transient stability.

3. SHORT-CIRCUIT STRENGTH (Ikss)
   Flag buses with Ikss < 5 kA as electrically weak (low fault-current
   infeed → poor damping, risk of voltage collapse after a fault).

4. BRANCH LOADING
   List every line or transformer with loading > 70 % (high stress) or
   > 100 % (overloaded).  Identify the most critical corridors.

5. NETWORK BOTTLENECKS
   Identify any single branch whose outage (N-1) would likely isolate a
   sub-network or cause severe overloads elsewhere.

6. GENERATION / LOAD BALANCE PER AREA
   Note buses with unusually large generation or load that may create
   local imbalance issues.

Be specific: quote exact bus names, branch names, and numerical values.
"""

_MITIGATION_PROMPT = """\
You are a senior transmission grid planning engineer responsible for stability enhancement and contingency mitigation design.

Your task is to propose technically sound mitigation measures based on:
- transient stability results
- grid behavior
- identified weak points

You MUST prioritize solutions based on effectiveness and feasibility in real power systems.

========================
INPUT DATA
========================
- Final report: {final_report}
- Numerical results: {numerics}
- Grid data: {grid_data}
- Study case: {study_case}

========================
CRITICAL RULES
========================
1. DO NOT propose generic textbook solutions without system-specific justification.
2. Every mitigation MUST reference a specific observed weakness (bus, generator, line, KPI).
3. Rank all actions by:
     - Effectiveness (stability improvement)
     - Implementation feasibility
     - Operational impact
4. Avoid vague suggestions like "improve control" or "optimize system".
5. Use power system engineering terminology (PSS, AVR tuning, fault clearing time, reactive compensation, line reinforcement, redispatch).
6. If grid data is insufficient, explicitly state limitations.
7. If the scenario is a generator disconnection (gen_switch open), do NOT treat that disconnected unit as oscillating; prioritize load shedding and placement.

========================
MANDATORY OUTPUT STRUCTURE
========================

## 1. System Weakness Diagnosis
Identify root causes:
- weakest buses
- overloaded lines
- poorly damped generators
- critical contingency points

Each must include numeric evidence.

## 2. Root Cause Analysis
Explain physical cause-effect chain:
Fault -> voltage dip -> power imbalance -> oscillations -> recovery behavior

Must be grounded in observed KPIs.

## 3. Ranked Mitigation Measures (MOST IMPORTANT SECTION)

Provide EXACTLY 5-10 actions ranked:

For each action:

### Action X - [Title]
- Type: (Protection / Control / Planning / Operational)
- Target component: (bus / line / generator)
- Observed issue addressed: (specific KPI)
- Mechanism: (how it improves stability)
- Expected impact: (quantified or qualitative but justified)
- Implementation complexity: Low / Medium / High

## 4. Priority Recommendation Set (Top 3)
- Must be actionable within real grid operations

## 5. Risk if No Action is Taken
Describe system vulnerability escalation (must be technical, not generic)

========================
OUTPUT STYLE
========================
- Transmission system planning report style
- Highly structured
- KPI-driven justification
- No generic advice
- No repetition of input data
"""


# ══════════════════════════════════════════════════════════════════
# INTERNAL HELPERS
# ══════════════════════════════════════════════════════════════════

def _nan(v) -> bool:
    try:
        return v is None or math.isnan(float(v))
    except (TypeError, ValueError):
        return True


def _fmt(v, fmt=".4f", unit="") -> str:
    if _nan(v):
        return "n/a"
    return f"{float(v):{fmt}}{unit}"


def _build_grid_text(grid_data: dict) -> str:
    """
    Convert the structured grid_data dict into a clear, annotated text block
    that an LLM can read as if it were looking at the topology graph.
    """
    buses      = grid_data.get("buses",      {})
    lines      = grid_data.get("lines",      [])
    loads      = grid_data.get("loads",      [])
    generators = grid_data.get("generators", [])

    sections: list[str] = []

    # ── Bus table ─────────────────────────────────────────────────
    sections.append("=== BUS DATA (load-flow + short-circuit) ===")
    sections.append(
        f"  {'Bus':20s}  {'V [p.u.]':>10}  {'Angle [deg]':>12}  {'Ikss [kA]':>10}"
    )
    for key in sorted(buses):
        info = buses[key]
        name  = info.get("name", key)
        v     = info.get("u")
        phi   = info.get("phiu")
        ikss  = info.get("ikss")
        flags = []
        if not _nan(v) and float(v) < 0.95:
            flags.append("LOW-V")
        if not _nan(v) and float(v) > 1.05:
            flags.append("HIGH-V")
        if not _nan(ikss) and float(ikss) < 5.0:
            flags.append("WEAK-Ikss")
        flag_str = "  *** " + ", ".join(flags) if flags else ""
        sections.append(
            f"  {name:20s}  {_fmt(v):>10}  {_fmt(phi, '.2f'):>12}  "
            f"{_fmt(ikss, '.2f'):>10}{flag_str}"
        )

    # ── Branch table ──────────────────────────────────────────────
    sections.append("")
    sections.append("=== BRANCH DATA (lines + transformers) ===")
    sections.append(
        f"  {'Branch':30s}  {'From':15s}  {'To':15s}  "
        f"{'Loading [%]':>11}  {'P_flow [MW]':>11}"
    )
    for item in sorted(lines, key=lambda x: -(float(x["loading"]) if not _nan(x["loading"]) else 0)):
        loading = item.get("loading")
        flags = ""
        if not _nan(loading):
            if float(loading) > 100:
                flags = "  *** OVERLOADED"
            elif float(loading) > 70:
                flags = "  *** HIGH"
        sections.append(
            f"  {item['name']:30s}  {item['bus1']:15s}  {item['bus2']:15s}  "
            f"{_fmt(loading, '.1f'):>11}  {_fmt(item.get('p_flow'), '.2f'):>11}{flags}"
        )

    # ── Generator table ───────────────────────────────────────────
    sections.append("")
    sections.append("=== GENERATORS ===")
    sections.append(f"  {'Generator':20s}  {'Bus':15s}  {'P [MW]':>8}  {'Q [MVAr]':>10}")
    for item in sorted(generators, key=lambda x: x["name"]):
        sections.append(
            f"  {item['name']:20s}  {item['bus']:15s}  "
            f"{_fmt(item.get('p_ini'), '.2f'):>8}  {_fmt(item.get('q_ini'), '.2f'):>10}"
        )

    # ── Load table ────────────────────────────────────────────────
    sections.append("")
    sections.append("=== LOADS ===")
    sections.append(f"  {'Load':20s}  {'Bus':15s}  {'P [MW]':>8}  {'Q [MVAr]':>10}")
    for item in sorted(loads, key=lambda x: x["name"]):
        sections.append(
            f"  {item['name']:20s}  {item['bus']:15s}  "
            f"{_fmt(item.get('p_ini'), '.2f'):>8}  {_fmt(item.get('q_ini'), '.2f'):>10}"
        )

    return "\n".join(sections)


def _build_kpi_summary(numerics: dict) -> str:
    """Compact fault-response KPI block used as ground truth for the LLM."""
    vd = numerics.get("voltages", {})
    sd = numerics.get("speeds",   {})
    ad = numerics.get("angles",   {})
    scenario = numerics.get("scenario", {})

    lines = [
        f"t_fault={numerics.get('t_fault')}s  "
        f"t_clear={numerics.get('t_clear')}s  "
        f"t_end={numerics.get('t_end')}s",
        (
            f"scenario: fault_type={scenario.get('fault_type', 'n/a')}  "
            f"fault_element={scenario.get('fault_element', 'n/a')}  "
            f"switch_element={scenario.get('switch_element', 'n/a')}  "
            f"switch_state={scenario.get('switch_state', 'n/a')}"
        ),
        "",
        "--- VOLTAGE RESPONSE (post-fault) ---",
    ]
    excluded = scenario.get("excluded_generators", [])
    if excluded:
        lines.append(
            "Excluded from speed/angle analysis (generator switched out): "
            + ", ".join(sorted(set(excluded)))
        )
        lines.append("")
    for bus, kpis in sorted(vd.items()):
        violation = "  [VIOLATION]" if kpis.get("out_of_limit_after_check") else ""
        lines.append(
            f"  {bus}: post_min={kpis['post_min']:.4f} p.u."
            f"  post_max={kpis['post_max']:.4f} p.u."
            f"  final={kpis['final_val']:.4f} p.u.{violation}"
        )

    lines += ["", "--- SPEED RESPONSE ---"]
    for gen, kpis in sorted(sd.items()):
        lines.append(
            f"  {gen}: min={kpis['min']:.5f}  max={kpis['max']:.5f}"
            f"  settle={kpis['settle_s']:.3f}s"
        )

    lines += ["", "--- ROTOR ANGLE SWING ---"]
    for gen, kpis in sorted(ad.items()):
        lines.append(
            f"  {gen}: delta_max={np.degrees(kpis['delta_max']):.2f} deg"
            f"  post_std={kpis['post_std']:.5f} rad"
        )

    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════
# AGENT 9 — MITIGATION AGENT
# ══════════════════════════════════════════════════════════════════

def mitigation_agent(final_report: str,
                     numerics: dict,
                     grid_data: dict,
                     study_case: str) -> str:
    """
    Analyse the pre-fault grid data and stability report, then propose
    ranked engineering mitigations.

    Parameters
    ----------
    final_report : str
        Text from Agent 6 (or Agent 4 if Agent 6 was skipped).
    numerics : dict
        KPI dict from analysis_agent (fault-response numerical data).
    grid_data : dict
        Structured pre-fault data from DIgSILENTAgent.export_grid_graph:
        buses, lines, loads, generators — used directly, no image needed.
    study_case : str
        Study-case name (used in section headings).

    Returns
    -------
    str
        Full mitigation report combining Stage A and Stage B outputs.
    """
    print("\n" + "═" * 60)
    print("  AGENT 9 — MITIGATION AGENT")
    print("═" * 60)

    grid_text = _build_grid_text(grid_data) if grid_data else "(Grid data unavailable)"

    # ── Stage A: identify vulnerabilities in the pre-fault grid ──
    print("[INFO] Stage A — analysing pre-fault grid data …")

    stage_a_msg = (
        f"Study case: {study_case}\n\n"
        "PRE-FAULT GRID DATA:\n"
        + grid_text
    )
    graph_analysis = run_agent(
        system_prompt=_GRAPH_ANALYSIS_PROMPT,
        user_message=stage_a_msg,
        max_tokens=1500,
        model=MODEL_SMART,
    )
    print("[OK] Stage A complete.")
    print(graph_analysis)

    # ── Stage B: propose mitigations ─────────────────────────────
    print("\n[INFO] Stage B — generating mitigation plan …")

    stage_b_msg = (
        f"Study case: {study_case}\n\n"
        "GRID TOPOLOGY ANALYSIS (Stage A):\n"
        + graph_analysis
        + "\n\nTRANSIENT STABILITY REPORT:\n"
        + final_report
        + "\n\nNUMERICAL KPIs (ground truth):\n"
        + _build_kpi_summary(numerics)
                + "\n\nIMPORTANT: Respect scenario logic exactly. If fault_type is gen_switch with switch_state=0, "
                    "the switched generator is out of service and must not be used as an active oscillating unit "
                    "for mitigation reasoning."
    )
    mitigation_plan = run_agent(
        system_prompt=_MITIGATION_PROMPT,
        user_message=stage_b_msg,
        max_tokens=3000,
        model=MODEL_SMART,
    )
    print("[OK] Stage B complete.")
    print(mitigation_plan)

    return (
        "═" * 60 + "\n"
        "  GRID TOPOLOGY ANALYSIS\n"
        + "═" * 60 + "\n"
        + graph_analysis
        + "\n\n"
        + "═" * 60 + "\n"
        "  MITIGATION PLAN\n"
        + "═" * 60 + "\n"
        + mitigation_plan
    )
