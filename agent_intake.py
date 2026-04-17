"""
Agent 0 — Intake Agent
=======================
Handles user input before any simulation is launched.  Supports three modes:

  1. Natural-language description  → LLM (Flash-Lite) parses it into a list
     of structured case-study dicts, enabling multi-case mode.
  2. Structured JSON / key:value   → parsed directly into SimulationConfig
     fields for a single-case run.
  3. Empty input                   → falls back to config-file defaults.

Public API
----------
  intake_agent(cfg)  →  (SimulationConfig | None,  list[dict] | None)

    Returns (updated_cfg, None)    for single-case mode.
    Returns (None, cases_list)     for multi-case mode.
"""

import json
import re

from Agent_DIgSILENT import SimulationConfig
from llm_client import MODEL_FAST, run_agent
from prompt_loader import get_prompt


def parse_natural_language_input(user_text: str) -> list[dict]:
    """
    Parse natural language descriptions and extract structured case study parameters.

    Returns a list of case studies, each with:
      {
        "case_name": str,
                "fault_type": "bus" | "line" | "gen_switch",
        "fault_element": str,
                "switch_element": str,
                "switch_state": int,
                "t_switch": float,
        "t_fault": float,
        "t_clear": float,
      }
    """
    print("\n" + "═" * 60)
    print("  AGENT 0 — INPUT PARSER (LLM)")
    print("═" * 60)

    system_prompt = get_prompt("agent_0_input_parser")

    user_msg = f"Extract case studies from this request:\n\n{user_text}"

    try:
        result = run_agent(
            system_prompt=system_prompt,
            user_message=user_msg,
            max_tokens=1500,
            model=MODEL_FAST,
        )

        cases = json.loads(result)
        if not isinstance(cases, list):
            cases = [cases]

        print(f"[OK] Parsed {len(cases)} case study(ies):")
        for case in cases:
            print(f"     • {case.get('case_name', 'Unknown')}: "
                  f"{case.get('fault_type')} @ {case.get('fault_element')}")
        return cases
    except json.JSONDecodeError as e:
        print(f"[ERROR] Failed to parse LLM response: {e}")
        print(f"        Response was: {result}")
        return []


def _coerce_user_request(raw: str) -> dict:
    """
    Parse user-provided configuration text.
    Supports JSON and loose key:value lines.
    """
    text = (raw or "").strip()
    if not text:
        return {}

    try:
        parsed = json.loads(text)
        return parsed if isinstance(parsed, dict) else {}
    except json.JSONDecodeError:
        pass

    wrapped = text
    if not wrapped.startswith("{"):
        wrapped = "{" + wrapped
    if not wrapped.endswith("}"):
        wrapped = wrapped + "}"
    try:
        parsed = json.loads(wrapped)
        return parsed if isinstance(parsed, dict) else {}
    except json.JSONDecodeError:
        pass

    out = {}
    for line in text.splitlines():
        line = line.strip().rstrip(",")
        if not line or ":" not in line:
            continue
        key, value = line.split(":", 1)
        key = key.strip().strip('"').strip("'")
        value = value.strip()
        if value.startswith('"') and value.endswith('"'):
            out[key] = value[1:-1]
            continue
        if value.startswith("'") and value.endswith("'"):
            out[key] = value[1:-1]
            continue
        if re.fullmatch(r"[-+]?\d+(\.\d+)?", value):
            out[key] = float(value)
            continue
        out[key] = value
    return out


def intake_agent(cfg: SimulationConfig) -> tuple[SimulationConfig | None, list[dict] | None]:
    """
    Enhanced intake agent that supports:
      1. Natural language input -> parse to multiple cases
      2. Structured JSON/key:value -> single case
      3. Empty input -> use defaults
    """
    print("\n" + "═" * 60)
    print("  AGENT 0 — INTAKE AGENT (ENHANCED)")
    print("═" * 60)
    print("\nOptions:")
    print("  1. Describe your simulation in NATURAL LANGUAGE")
    print("     (e.g., 'I want to simulate two case studies...')")
    print("  2. Provide structured JSON or key:value pairs (legacy)")
    print("  3. Press Enter to keep defaults")
    print("\nDescribe what simulation(s) you want to run:")
    print("(For multi-case scenarios, end with an empty line)\n")

    user_lines = []
    while True:
        line = input("> ")
        if not line.strip():
            break
        user_lines.append(line)

    user_raw = "\n".join(user_lines).strip()
    if not user_raw:
        print("[INFO] No user request provided. Using config defaults (single case).")
        return cfg, None

    is_natural_language = (
        ("case" in user_raw.lower() or "scenario" in user_raw.lower()) and
        ("simulate" in user_raw.lower() or "run" in user_raw.lower())
    )
    is_structured = ("fault_type" in user_raw or "fault_element" in user_raw
                     or user_raw.strip().startswith("{"))

    if is_natural_language and not is_structured:
        cases = parse_natural_language_input(user_raw)
        if cases and len(cases) > 0:
            print(f"\n[OK] Parsed {len(cases)} case(s) for multi-case pipeline")
            return None, cases
        print("[WARN] Failed to parse natural language input.")
        return cfg, None

    req = _coerce_user_request(user_raw)
    if not req:
        print("[WARN] Could not parse request. Using config defaults.")
        return cfg, None

    orig_t_fault = cfg.t_fault
    orig_t_clear = cfg.t_clear

    if "fault_type" in req:
        cfg.fault_type = str(req["fault_type"]).strip().lower()
    if "fault_element" in req:
        cfg.fault_element = str(req["fault_element"]).strip()
    if "t_fault" in req:
        cfg.t_fault = float(req["t_fault"])
    if "t_clear" in req:
        cfg.t_clear = float(req["t_clear"])
    if "t_end" in req:
        cfg.t_end = float(req["t_end"])
    if "dt_rms" in req:
        cfg.dt_rms = float(req["dt_rms"])
    if "run_label" in req:
        cfg.run_label = str(req["run_label"]).strip()
    if "study_case" in req:
        cfg.study_case = str(req["study_case"]).strip()

    if cfg.t_clear <= cfg.t_fault:
        cfg.t_fault = orig_t_fault
        cfg.t_clear = orig_t_clear
        print("[WARN] t_clear must be greater than t_fault. Reverted to original timing values.")

    print(
        "[OK] Intake applied (single case): "
        f"fault_type={cfg.fault_type}, "
        f"fault_element={cfg.fault_element}, "
        f"t_fault={cfg.t_fault}, t_clear={cfg.t_clear}"
    )
    return cfg, None
