"""
Agent 3 — Plot Agent
=====================
Generates time-series PNG plots from parsed RMS simulation data and saves
them to the run output directory.

Single-case plots (plot_agent)
------------------------------
  *_voltages.png            — all bus voltage magnitudes [p.u.]
  *_speeds.png              — all generator rotor speeds [p.u.]
  *_<bus>_voltage_violation.png — individual plots for buses that violate
                                  the ±10 % voltage band after clearing

Multi-case comparison plots
----------------------------
  plot_voltage_case_comparison  — voltage time-series overlay per violated bus
  plot_speed_case_comparison    — speed time-series overlay for top-N generators
                                  by worst |speed − 1| deviation

All figures use matplotlib's non-interactive Agg backend so no display is
required (suitable for headless / server execution).
"""

import os
import re

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

from Agent_DIgSILENT import SimulationConfig


def _excluded_generators_from_numerics(numerics: dict) -> set[str]:
    scenario = numerics.get("scenario", {}) or {}
    excluded = scenario.get("excluded_generators", []) or []
    return {str(item).strip().lower() for item in excluded if str(item).strip()}


def plot_agent(numerics: dict, cfg: SimulationConfig) -> list[str]:
    print("\n" + "═" * 60)
    print("  AGENT 3 — PLOT AGENT")
    print("═" * 60)

    parsed = numerics.get("_parsed", {})
    time_ = parsed.get("time", np.array([]))
    signals = parsed.get("signals", {})
    t_clear = numerics["t_clear"]
    dt_rms = float(getattr(cfg, "dt_rms", numerics.get("dt", 0.0)))
    violation_time = t_clear + 2.0 * dt_rms
    out_dir = cfg.output_dir
    label = cfg.run_label
    saved = []

    if len(time_) == 0:
        print("[WARN] No data to plot.")
        return saved

    groups: dict[str, dict[str, np.ndarray]] = {}
    for (obj, var), arr in signals.items():
        groups.setdefault(var, {})[obj] = arr

    volt_key = "u1, Magnitude in p.u."
    if volt_key in groups:
        fig, ax = plt.subplots(figsize=(12, 5))
        for obj, arr in sorted(groups[volt_key].items()):
            ax.plot(time_, arr, lw=0.9, label=obj)
        ax.axhline(0.9, color="gray", lw=0.8, ls="--")
        ax.axhline(1.1, color="gray", lw=0.8, ls="--")
        ax.set_xlabel("Time [s]")
        ax.set_ylabel("Voltage [p.u.]")
        ax.set_title(f"Bus Voltage Magnitudes — {cfg.study_case}")
        ax.legend(fontsize=6, ncol=4, loc="lower right")
        ax.grid(True, alpha=0.3)
        path = os.path.join(out_dir, f"{label}_voltages.png")
        fig.tight_layout()
        fig.savefig(path, dpi=150)
        plt.close(fig)
        saved.append(path)
        print(f"[OK] Saved -> {path}")

    spd_key = "Speed in p.u."
    if spd_key in groups:
        excluded_generators = _excluded_generators_from_numerics(numerics)
        fig, ax = plt.subplots(figsize=(12, 5))
        for obj, arr in sorted(groups[spd_key].items()):
            if str(obj).strip().lower() in excluded_generators:
                continue
            ax.plot(time_, arr, lw=0.9, label=obj)
        ax.axhline(1.0, color="black", lw=0.8, ls="-", alpha=0.4)
        ax.axhline(1.005, color="gray", lw=0.7, ls=":")
        ax.axhline(0.995, color="gray", lw=0.7, ls=":")
        ax.set_xlabel("Time [s]")
        ax.set_ylabel("Speed [p.u.]")
        ax.set_title(f"Generator Speeds — {cfg.study_case}")
        ax.legend(fontsize=7, ncol=3, loc="lower right")
        ax.grid(True, alpha=0.3)
        path = os.path.join(out_dir, f"{label}_speeds.png")
        fig.tight_layout()
        fig.savefig(path, dpi=150)
        plt.close(fig)
        saved.append(path)
        print(f"[OK] Saved -> {path}")

    violated_buses = []
    if volt_key in groups:
        for obj, arr in sorted(groups[volt_key].items()):
            post_arr = arr[time_ >= violation_time]
            if post_arr.size == 0:
                continue
            if np.any((post_arr < 0.9) | (post_arr > 1.1)):
                violated_buses.append((obj, arr))

    for obj, arr in violated_buses:
        fig, ax = plt.subplots(figsize=(12, 5))
        ax.plot(time_, arr, lw=1.0, label=obj)
        # ax.axvline(violation_time, color="purple", lw=0.8, ls=":", label="Check time")
        ax.axhline(0.9, color="gray", lw=0.8, ls="--")
        ax.axhline(1.1, color="gray", lw=0.8, ls="--")
        ax.set_xlabel("Time [s]")
        ax.set_ylabel("Voltage [p.u.]")
        ax.set_title(f"{obj} Voltage — Out of Limit")
        ax.legend(fontsize=7, loc="best")
        ax.grid(True, alpha=0.3)
        safe_obj = re.sub(r"[^A-Za-z0-9_.-]+", "_", obj)
        path = os.path.join(out_dir, f"{label}_{safe_obj}_voltage_violation.png")
        fig.tight_layout()
        fig.savefig(path, dpi=150)
        plt.close(fig)
        saved.append(path)
        print(f"[OK] Saved -> {path}")

    return saved


def plot_voltage_case_comparison(case_results: dict, out_dir: str, label: str = "comparison") -> list[str]:
    print("\n" + "═" * 60)
    print("  AGENT 3 — VOLTAGE CASE COMPARISON")
    print("═" * 60)

    case_names = list(case_results.keys())
    if len(case_names) < 2:
        print("[WARN] At least 2 cases needed for voltage comparison.")
        return []

    violated_buses: set[str] = set()
    for case_name in case_names:
        numerics = case_results.get(case_name, {}).get("numerics", {})
        for bus, kpis in numerics.get("voltages", {}).items():
            if kpis.get("out_of_limit_after_check"):
                violated_buses.add(bus)

    if not violated_buses:
        print("[INFO] No bus violated voltage limits across the compared cases.")
        return []

    saved: list[str] = []
    for bus in sorted(violated_buses):
        fig, ax = plt.subplots(figsize=(12, 5))
        for case_name in case_names:
            parsed = case_results.get(case_name, {}).get("parsed", {})
            time_ = parsed.get("time", np.array([]))
            signals = parsed.get("signals", {})
            arr = signals.get((bus, "u1, Magnitude in p.u."))
            if arr is None or len(time_) == 0:
                continue
            ax.plot(time_, arr, lw=1.0, label=case_name)

        ax.axhline(0.9, color="gray", lw=0.8, ls="--")
        ax.axhline(1.1, color="gray", lw=0.8, ls="--")
        ax.set_xlabel("Time [s]")
        ax.set_ylabel("Voltage [p.u.]")
        ax.set_title(f"Voltage Time-Series Comparison for {bus}")
        ax.legend(fontsize=7)
        ax.grid(True, alpha=0.3)
        path = os.path.join(out_dir, f"{label}_{re.sub(r'[^A-Za-z0-9_.-]+', '_', bus)}_voltage_comparison.png")
        fig.tight_layout()
        fig.savefig(path, dpi=150)
        plt.close(fig)
        saved.append(path)
        print(f"[OK] Saved -> {path}")

    return saved


def plot_speed_case_comparison(case_results: dict,
                               out_dir: str,
                               label: str = "comparison",
                               top_n: int = 3) -> list[str]:
    print("\n" + "═" * 60)
    print("  AGENT 3 — SPEED CASE COMPARISON")
    print("═" * 60)

    case_names = list(case_results.keys())
    if len(case_names) < 2:
        print("[WARN] At least 2 cases needed for speed comparison.")
        return []

    # Collect worst absolute speed deviation from 1.0 p.u. across all cases per generator.
    severity_by_gen: dict[str, float] = {}
    for case_name in case_names:
        parsed = case_results.get(case_name, {}).get("parsed", {})
        numerics = case_results.get(case_name, {}).get("numerics", {})
        excluded_generators = _excluded_generators_from_numerics(numerics)
        signals = parsed.get("signals", {})
        for (obj, var), arr in signals.items():
            if var != "Speed in p.u." or arr is None or len(arr) == 0:
                continue
            if str(obj).strip().lower() in excluded_generators:
                continue
            deviation = float(np.max(np.abs(np.asarray(arr) - 1.0)))
            prev = severity_by_gen.get(obj, 0.0)
            if deviation > prev:
                severity_by_gen[obj] = deviation

    if not severity_by_gen:
        print("[INFO] No generator speed signals found across compared cases.")
        return []

    n = max(1, int(top_n))
    critical_gens = sorted(severity_by_gen, key=lambda g: severity_by_gen[g], reverse=True)[:n]
    print(
        "[INFO] Critical generators by worst |speed-1|: "
        + ", ".join(f"{g} ({severity_by_gen[g]:.5f})" for g in critical_gens)
    )

    saved: list[str] = []
    for gen in critical_gens:
        fig, ax = plt.subplots(figsize=(12, 5))
        plotted = 0
        for case_name in case_names:
            parsed = case_results.get(case_name, {}).get("parsed", {})
            numerics = case_results.get(case_name, {}).get("numerics", {})
            excluded_generators = _excluded_generators_from_numerics(numerics)
            time_ = parsed.get("time", np.array([]))
            signals = parsed.get("signals", {})
            arr = signals.get((gen, "Speed in p.u."))
            if arr is None or len(time_) == 0:
                continue
            if str(gen).strip().lower() in excluded_generators:
                continue
            ax.plot(time_, arr, lw=1.0, label=case_name)
            plotted += 1

        if plotted == 0:
            plt.close(fig)
            continue

        ax.axhline(1.0, color="black", lw=0.8, ls="-", alpha=0.4)
        ax.axhline(1.005, color="gray", lw=0.7, ls=":")
        ax.axhline(0.995, color="gray", lw=0.7, ls=":")
        ax.set_xlabel("Time [s]")
        ax.set_ylabel("Speed [p.u.]")
        ax.set_title(f"Speed Time-Series Comparison for {gen}")
        ax.legend(fontsize=7)
        ax.grid(True, alpha=0.3)
        safe_gen = re.sub(r"[^A-Za-z0-9_.-]+", "_", gen)
        path = os.path.join(out_dir, f"{label}_{safe_gen}_speed_comparison.png")
        fig.tight_layout()
        fig.savefig(path, dpi=150)
        plt.close(fig)
        saved.append(path)
        print(f"[OK] Saved -> {path}")

    return saved
