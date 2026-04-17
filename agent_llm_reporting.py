"""
Agents 4–7 — LLM Reporting Agents
===================================
Four LLM-powered agents that transform numerical KPIs into engineering reports.

Agent 4 — summary_agent
    Writes the first narrative draft: voltage stability, rotor-angle stability,
    speed stability, and an overall verdict.  Uses MODEL_FAST (low latency / cost).

Agent 5 — review_agent
    Cross-checks every numerical claim in the draft against the raw KPIs.
    Returns a numbered correction list, or "No improvements needed."
    Uses MODEL_FAST.

Agent 6 — final_report_agent
    Applies all reviewer corrections and produces the polished final report
    ready for publication.  Uses MODEL_SMART for higher quality output.

Agent 7 — comparison_agent
    Compares KPIs across multiple case studies, producing a structured
    comparative risk assessment and recommendations.  Uses MODEL_SMART.

All system prompts are loaded from agent_prompts.csv via prompt_loader so
they can be tuned without touching this file.
"""

import time

import numpy as np

from llm_client import MODEL_FAST, MODEL_SMART, run_agent
from prompt_loader import get_prompt


def _build_kpi_block(numerics: dict,
                     compact: bool = False,
                     max_voltage: int = 12,
                     max_speed: int = 8,
                     max_angle: int = 8) -> str:
    vd = numerics.get("voltages", {})
    sd = numerics.get("speeds", {})
    ad = numerics.get("angles", {})
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
        "--- VOLTAGE KPIs (post-fault, all buses) ---",
    ]
    excluded = scenario.get("excluded_generators", [])
    if excluded:
        lines.append(
            "--- NOTE --- Excluded from speed/angle assessment due to gen_switch-open event: "
            + ", ".join(sorted(set(excluded)))
        )
        lines.append("")
    voltage_items = sorted(vd.items())
    if compact and voltage_items:
        def _voltage_severity(item):
            _, k = item
            vmin = float(k.get("post_min", np.nan))
            vmax = float(k.get("post_max", np.nan))
            low_dev = max(0.0, 0.9 - vmin) if not np.isnan(vmin) else 0.0
            high_dev = max(0.0, vmax - 1.1) if not np.isnan(vmax) else 0.0
            vio = 1.0 if bool(k.get("out_of_limit_after_check")) else 0.0
            return (vio, max(low_dev, high_dev))
        voltage_items = sorted(voltage_items, key=_voltage_severity, reverse=True)[:max_voltage]

    for bus, kpis in voltage_items:
        lines.append(
            f"  {bus}: post_min={kpis['post_min']:.4f} p.u. at t={kpis['t_post_min']:.3f}s  "
            f"post_max={kpis['post_max']:.4f} p.u. at t={kpis['t_post_max']:.3f}s  "
            f"final={kpis['final_val']:.4f} p.u."
        )
    if compact and len(vd) > len(voltage_items):
        lines.append(f"  ... {len(vd) - len(voltage_items)} additional buses omitted for compact mode")
    lines.append("")
    lines.append("--- SPEED KPIs (all generators) ---")
    speed_items = sorted(sd.items())
    if compact and speed_items:
        def _speed_severity(item):
            _, k = item
            return max(abs(float(k.get("min", 1.0)) - 1.0), abs(float(k.get("max", 1.0)) - 1.0))
        speed_items = sorted(speed_items, key=_speed_severity, reverse=True)[:max_speed]

    for gen, kpis in speed_items:
        lines.append(
            f"  {gen}: min={kpis['min']:.5f}  max={kpis['max']:.5f}  "
            f"post_mean={kpis['post_mean']:.5f}  settle={kpis['settle_s']:.3f}s"
        )
    if compact and len(sd) > len(speed_items):
        lines.append(f"  ... {len(sd) - len(speed_items)} additional generators omitted for compact mode")
    lines.append("")
    lines.append("--- ROTOR ANGLE KPIs (all generators) ---")
    angle_items = sorted(ad.items())
    if compact and angle_items:
        def _angle_severity(item):
            _, k = item
            return abs(float(k.get("delta_span", k.get("delta_max", 0.0))))
        angle_items = sorted(angle_items, key=_angle_severity, reverse=True)[:max_angle]

    for gen, kpis in angle_items:
        lines.append(
            f"  {gen}: delta_max={np.degrees(kpis['delta_max']):.2f}°  "
            f"post_std={kpis['post_std']:.5f} rad"
        )
        if "angle_t0" in kpis or "angle_t_end" in kpis or "delta_span" in kpis:
            lines.append(
                f"    angle_t0={kpis.get('angle_t0', float('nan')):.5f} rad  "
                f"angle_t_end={kpis.get('angle_t_end', float('nan')):.5f} rad  "
                f"delta_span={kpis.get('delta_span', float('nan')):.5f} rad"
            )
    if compact and len(ad) > len(angle_items):
        lines.append(f"  ... {len(ad) - len(angle_items)} additional generators omitted for compact mode")

    lines.append("")
    lines.append("--- VOLTAGE SNAPSHOT KPIs ---")
    for bus, kpis in voltage_items:
        lines.append(
            f"  {bus}: v_t0={kpis.get('v_t0', float('nan')):.4f} p.u.  "
            f"v_t_end={kpis.get('v_t_end', float('nan')):.4f} p.u.  "
            f"post_clear_min={kpis.get('post_clear_min', float('nan')):.4f} p.u.  "
            f"post_clear_max={kpis.get('post_clear_max', float('nan')):.4f} p.u."
        )
    return "\n".join(lines)


def _build_voltage_case_comparison_block(results_dict: dict, case_names: list[str]) -> str:
    violated_buses: set[str] = set()
    for case_name in case_names:
        case_numerics = results_dict.get(case_name, {})
        for bus, kpis in case_numerics.get("voltages", {}).items():
            if kpis.get("out_of_limit_after_check"):
                violated_buses.add(bus)

    if not violated_buses:
        return "--- VOLTAGE COMPARISON ---\n  No buses were out of limits after clearing + 2 dt across the compared cases."

    lines = ["--- VOLTAGE COMPARISON (buses with at least one violation) ---"]
    for bus in sorted(violated_buses):
        lines.append(f"  {bus}")
        for case_name in case_names:
            kpis = results_dict.get(case_name, {}).get("voltages", {}).get(bus)
            if not kpis:
                continue
            violation = "YES" if kpis.get("out_of_limit_after_check") else "NO"
            lines.append(
                f"    {case_name}: post_min={kpis.get('post_min', float('nan')):.4f} p.u., "
                f"post_max={kpis.get('post_max', float('nan')):.4f} p.u., "
                f"final={kpis.get('final_val', float('nan')):.4f} p.u., violation={violation}"
            )
    return "\n".join(lines)


def summary_agent(numerics: dict, study_case: str) -> str:
    print("\n" + "═" * 60)
    print("  AGENT 4 — SUMMARY AGENT (LLM)")
    print("═" * 60)

    kpi_block = _build_kpi_block(numerics, compact=True)
    user_msg = (
        f"Study case: {study_case}\n\n"
        "NUMERICAL KPIs (ground truth):\n"
        + kpi_block
    )

    summary = run_agent(
        system_prompt=get_prompt("agent_4_summary"),
        user_message=user_msg,
        max_tokens=2000,
        model=MODEL_FAST,
    )

    print(summary)
    return summary


def review_agent(numerics: dict, summary: str) -> str:
    print("\n" + "═" * 60)
    print("  AGENT 5 — REVIEW AGENT")
    print("═" * 60)

    user_msg = (
        "NUMERICAL KPIs (ground truth):\n"
        + _build_kpi_block(numerics, compact=True)
        + "\n\nDRAFT REPORT TO REVIEW:\n"
        + summary
    )

    improvements = run_agent(
        system_prompt=get_prompt("agent_5_review"),
        user_message=user_msg,
        max_tokens=1000,
        model=MODEL_FAST,
    )

    print(improvements)
    return improvements


def final_report_agent(summary: str,
                       improvements: str, study_case: str) -> str:
    print("\n" + "═" * 60)
    print("  AGENT 6 — FINAL REPORT AGENT")
    print("═" * 60)

    print("[INFO] Waiting 5s before final report call (rate limit buffer)...")
    time.sleep(5)

    user_msg = (
        f"Study case: {study_case}\n\n"
        "DRAFT REPORT:\n"
        + summary
        + "\n\nREVIEW CORRECTIONS:\n"
        + improvements
    )

    final_report = run_agent(
        system_prompt=get_prompt("agent_6_final_report"),
        user_message=user_msg,
        max_tokens=3000,
        model=MODEL_SMART,
    )

    print(final_report)
    return final_report


def comparison_agent(results_dict: dict, case_names: list[str]) -> str:
    print("\n" + "═" * 60)
    print("  AGENT 7 — COMPARISON AGENT (LLM)")
    print("═" * 60)

    if len(case_names) < 2:
        print("[WARN] At least 2 cases needed for comparison. Skipping.")
        return ""

    kpi_blocks = {}
    for case_name in case_names:
        if case_name in results_dict:
            kpi_blocks[case_name] = _build_kpi_block(results_dict[case_name], compact=True)

    voltage_comparison_block = _build_voltage_case_comparison_block(results_dict, case_names)

    comparison_text = "CASE STUDIES COMPARISON\n"
    comparison_text += "═" * 60 + "\n\n"

    for case_name in case_names:
        if case_name in kpi_blocks:
            comparison_text += f"--- {case_name} ---\n"
            comparison_text += kpi_blocks[case_name] + "\n\n"

    comparison_text += voltage_comparison_block + "\n\n"

    user_msg = comparison_text

    print("[INFO] Waiting 3s before comparison analysis (rate limit buffer)...")
    time.sleep(3)

    comparison_report = run_agent(
        system_prompt=get_prompt("agent_7_comparison"),
        user_message=user_msg,
        max_tokens=3000,
        model=MODEL_SMART,
    )

    print(comparison_report)
    return comparison_report
