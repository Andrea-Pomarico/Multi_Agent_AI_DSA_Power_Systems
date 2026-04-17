"""
Agent 2 — Analysis Agent
=========================
Parses the PowerFactory ComRes CSV and computes structured numerical KPIs
for use by all downstream agents (plot, LLM reporting, comparison).

CSV format expected
-------------------
  Row 0 : object names  (e.g. "Bus 01.ElmTerm")
  Row 1 : variable names (e.g. "u1, Magnitude in p.u.")
  Row 2+: numeric data rows, semicolon-separated, comma as decimal separator

Signals extracted
-----------------
  u1, Magnitude in p.u.   — bus voltage magnitudes
  Speed in p.u.            — generator rotor speeds
  phi in rad               — generator rotor angles

KPIs computed (per signal per object)
--------------------------------------
  Voltages : pre-fault mean, fault nadir, post-fault min/max/final/std,
             settling time (±5 % band around pre-fault mean)
  Speeds   : pre/post mean, global min/max, settling time (±0.5 % of 1 pu)
  Angles   : pre-fault mean, global min/max, max swing Δ, post-fault std

Public API
----------
  analysis_agent(csv_path, t_fault, t_clear)  →  (numerics: dict, stats_txt: str)
"""

import csv as csv_mod

import numpy as np


def _normalize_pf_name(name: str) -> tuple[str, str]:
    """Return normalized full name and base name (without class suffix)."""
    text = str(name or "").strip().strip('"').lower()
    base = text.split(".", 1)[0].strip()
    return text, base


def _is_switched_out_generator(obj_name: str,
                               fault_type: str | None,
                               switch_state: int | None,
                               switch_element: str | None,
                               fault_element: str | None) -> bool:
    """
    True when this object is the generator intentionally switched open
    in a gen_switch scenario.
    """
    ft = str(fault_type or "").strip().lower().replace("-", "_").replace(" ", "_")
    if ft in ("generator", "switch", "generator_switch"):
        ft = "gen_switch"
    if ft != "gen_switch":
        return False

    try:
        state = int(switch_state) if switch_state is not None else 0
    except (TypeError, ValueError):
        state = 0
    if state != 0:
        return False

    target = str(switch_element or fault_element or "").strip()
    if not target:
        return False

    obj_full, obj_base = _normalize_pf_name(obj_name)
    tgt_full, tgt_base = _normalize_pf_name(target)
    return (obj_full == tgt_full) or (obj_base and obj_base == tgt_base)


def _parse_csv(csv_path: str) -> dict:
    """
    Parse the PowerFactory ComRes CSV (semicolon-separated, 2 header rows).
    Returns:
      {
        "time": np.array,
        "signals": { (obj_name, variable_label): np.array, ... }
      }
    """
    with open(csv_path, "r", encoding="utf-8") as f:
        reader = csv_mod.reader(f, delimiter=";")
        rows = list(reader)

    obj_names = rows[0]
    var_names = rows[1]
    data_rows = rows[2:]
    n_cols = len(obj_names)

    mat = []
    for row in data_rows:
        try:
            mat.append([float(v.replace(",", ".")) for v in row[:n_cols]])
        except ValueError:
            continue
    mat = np.array(mat)

    if mat.shape[0] == 0:
        return {"time": np.array([]), "signals": {}}

    keep_vars = {
        "u1, Magnitude in p.u.",
        "Speed in p.u.",
        "phi in rad",
    }

    signals = {}
    for col_idx in range(1, n_cols):
        obj = obj_names[col_idx].strip().strip('"')
        var = var_names[col_idx].strip().strip('"')
        if var in keep_vars:
            signals[(obj, var)] = mat[:, col_idx]

    return {"time": mat[:, 0], "signals": signals}


def _settling_time(time: np.ndarray, arr: np.ndarray,
                   ref: float, band: float, t_start: float) -> float:
    """
    First time after t_start at which the signal stays within ref +- band
    for all subsequent samples. Returns nan if never settled.
    """
    mask = time >= t_start
    t_post = time[mask]
    s_post = arr[mask]
    for i in range(len(t_post)):
        if np.all(np.abs(s_post[i:] - ref) <= band):
            return float(t_post[i])
    return float("nan")


def analysis_agent(csv_path: str,
                   t_fault: float,
                   t_clear: float,
                   fault_type: str = "bus",
                   switch_element: str = "",
                   fault_element: str = "",
                   switch_state: int = 0,
                   t_switch: float | None = None) -> tuple[dict, str]:
    """
    Returns:
      numerics  - dict of structured KPIs
      stats_txt - formatted text for the LLM prompt
    """
    print("\n" + "═" * 60)
    print("  AGENT 2 — ANALYSIS AGENT")
    print("═" * 60)

    parsed = _parse_csv(csv_path)
    time_ = parsed["time"]
    signals = parsed["signals"]

    if len(time_) == 0 or not signals:
        print("[WARN] No data found in CSV.")
        return {}, "No data available."

    pre_mask = time_ < t_fault
    dur_mask = (time_ >= t_fault) & (time_ <= t_clear + (time_[1] - time_[0]))
    post_mask = time_ > t_clear + (time_[1] - time_[0])

    groups: dict[str, dict[str, np.ndarray]] = {}
    for (obj, var), arr in signals.items():
        groups.setdefault(var, {})[obj] = arr

    numerics: dict = {
        "t_fault": t_fault,
        "t_clear": t_clear,
        "t_end": float(time_[-1]),
        "dt": float(time_[1] - time_[0]) if len(time_) > 1 else 0.0,
        "check_time": float(t_clear + 2.0 * (time_[1] - time_[0])) if len(time_) > 1 else float(t_clear),
        "voltages": {},
        "speeds": {},
        "angles": {},
        "scenario": {
            "fault_type": str(fault_type or "").strip().lower(),
            "fault_element": str(fault_element or "").strip(),
            "switch_element": str(switch_element or "").strip(),
            "switch_state": int(switch_state) if str(switch_state).strip() else 0,
            "t_switch": float(t_switch) if t_switch is not None else float(t_fault),
            "excluded_generators": [],
        },
    }

    post_clear_start = t_clear + (3.0 * numerics["dt"])

    lines = [
        f"RMS Simulation  |  t_fault={t_fault}s  t_clear={t_clear}s  "
        f"t_end={time_[-1]:.2f}s  dt={numerics['dt']:.4f}s",
        "",
    ]

    volt_key = "u1, Magnitude in p.u."
    if volt_key in groups:
        lines.append("VOLTAGE MAGNITUDES [p.u.]")
        lines.append(f"  {'Bus':10s}  {'V(t0)':>8}  {'V(t_end)':>8}  "
                     f"{'PostClrMin':>10}  {'t_min(s)':>9}  {'PostClrMax':>10}  {'t_max(s)':>9}")
        for obj, arr in sorted(groups[volt_key].items()):
            out_of_limit_after_check = False
            pre_mean = float(np.mean(arr[pre_mask])) if pre_mask.any() else float("nan")
            nadir = float(np.min(arr[dur_mask])) if dur_mask.any() else float("nan")
            post_mean = float(np.mean(arr[post_mask])) if post_mask.any() else float("nan")
            post_std = float(np.std(arr[post_mask])) if post_mask.any() else float("nan")
            settle = _settling_time(time_, arr, pre_mean, 0.05, t_clear)
            v_t0 = float(arr[0])
            v_t_end = float(arr[-1])
            post_clear_mask = time_ >= post_clear_start
            if post_clear_mask.any():
                post_clear_arr = arr[post_clear_mask]
                post_clear_time = time_[post_clear_mask]
                post_clear_min = float(np.min(post_clear_arr))
                t_post_clear_min = float(post_clear_time[np.argmin(post_clear_arr)])
                post_clear_max = float(np.max(post_clear_arr))
                t_post_clear_max = float(post_clear_time[np.argmax(post_clear_arr)])
            else:
                post_clear_min = t_post_clear_min = post_clear_max = t_post_clear_max = float("nan")
            if post_mask.any():
                post_arr = arr[post_mask]
                post_time = time_[post_mask]
                post_min = float(np.min(post_arr))
                t_post_min = float(post_time[np.argmin(post_arr)])
                post_max = float(np.max(post_arr))
                t_post_max = float(post_time[np.argmax(post_arr)])
                final_val = float(post_arr[-1])
            else:
                post_min = t_post_min = post_max = t_post_max = final_val = float("nan")
            if len(time_) > 1:
                check_mask = time_ >= numerics["check_time"]
                if check_mask.any():
                    check_arr = arr[check_mask]
                    out_of_limit_after_check = bool(np.any((check_arr < 0.9) | (check_arr > 1.1)))
            numerics["voltages"][obj] = dict(
                pre_mean=pre_mean,
                nadir=nadir,
                post_mean=post_mean,
                post_std=post_std,
                settle_s=settle,
                v_t0=v_t0,
                v_t_end=v_t_end,
                post_clear_min=post_clear_min,
                t_post_clear_min=t_post_clear_min,
                post_clear_max=post_clear_max,
                t_post_clear_max=t_post_clear_max,
                post_min=post_min,
                t_post_min=t_post_min,
                post_max=post_max,
                t_post_max=t_post_max,
                final_val=final_val,
                out_of_limit_after_check=out_of_limit_after_check,
            )
            lines.append(
                f"  {obj:10s}  {v_t0:8.4f}  {v_t_end:8.4f}  "
                f"{post_clear_min:10.4f}  {t_post_clear_min:9.3f}  "
                f"{post_clear_max:10.4f}  {t_post_clear_max:9.3f}"
            )
        lines.append("")

    spd_key = "Speed in p.u."
    if spd_key in groups:
        lines.append("GENERATOR SPEEDS [p.u.]")
        lines.append(f"  {'Gen':10s}  {'Pre(mean)':>10}  {'Min':>8}  {'Max':>8}  "
                     f"{'Post(mean)':>10}  {'Post(std)':>9}  {'Settle(s)':>9}")
        for obj, arr in sorted(groups[spd_key].items()):
            if _is_switched_out_generator(
                obj,
                fault_type=fault_type,
                switch_state=switch_state,
                switch_element=switch_element,
                fault_element=fault_element,
            ):
                numerics["scenario"]["excluded_generators"].append(obj)
                continue
            pre_mean = float(np.mean(arr[pre_mask])) if pre_mask.any() else float("nan")
            spd_min = float(np.min(arr))
            spd_max = float(np.max(arr))
            post_mean = float(np.mean(arr[post_mask])) if post_mask.any() else float("nan")
            post_std = float(np.std(arr[post_mask])) if post_mask.any() else float("nan")
            settle = _settling_time(time_, arr, 1.0, 0.005, t_clear)
            numerics["speeds"][obj] = dict(
                pre_mean=pre_mean,
                min=spd_min,
                max=spd_max,
                post_mean=post_mean,
                post_std=post_std,
                settle_s=settle,
            )
            lines.append(
                f"  {obj:10s}  {pre_mean:10.4f}  {spd_min:8.4f}  {spd_max:8.4f}  "
                f"{post_mean:10.4f}  {post_std:9.5f}  {settle:9.3f}"
            )
        lines.append("")

    ang_key = "phi in rad"
    if ang_key in groups:
        lines.append("ROTOR ANGLES [rad]")
        lines.append(f"  {'Gen':10s}  {'Angle(t0)':>10}  {'Angle(t_end)':>12}  "
                     f"{'Δmin(rad)':>10}  {'Δmax(rad)':>10}  {'Δspan(rad)':>10}")
        for obj, arr in sorted(groups[ang_key].items()):
            if obj in numerics["scenario"]["excluded_generators"]:
                continue
            pre_mean = float(np.mean(arr[pre_mask])) if pre_mask.any() else float("nan")
            ang_min = float(np.min(arr))
            ang_max = float(np.max(arr))
            delta_max = ang_max - pre_mean
            angle_t0 = float(arr[0])
            angle_t_end = float(arr[-1])
            delta_span = ang_max - ang_min
            post_std = float(np.std(arr[post_mask])) if post_mask.any() else float("nan")
            numerics["angles"][obj] = dict(
                pre_mean=pre_mean,
                angle_t0=angle_t0,
                angle_t_end=angle_t_end,
                min=ang_min,
                max=ang_max,
                delta_min=ang_min,
                delta_max=ang_max,
                delta_span=delta_span,
                delta_rel_pre=delta_max,
                post_std=post_std,
            )
            lines.append(
                f"  {obj:10s}  {angle_t0:10.4f}  {angle_t_end:12.4f}  "
                f"{ang_min:10.4f}  {ang_max:10.4f}  {delta_span:10.4f}"
            )
        lines.append("")

    excluded = numerics.get("scenario", {}).get("excluded_generators", [])
    if excluded:
        lines.append(
            "Excluded from dynamic speed/angle analysis (gen_switch open event): "
            + ", ".join(sorted(set(excluded)))
        )
        lines.append("")

    stats_txt = "\n".join(lines)
    print(stats_txt)
    print(f"[OK] Analysed {len(signals)} signals over {len(time_)} time steps")

    numerics["_parsed"] = parsed
    return numerics, stats_txt
