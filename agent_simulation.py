"""
Agent 1 — Simulation Agent
===========================
Thin wrapper around DIgSILENTAgent that runs the full PowerFactory RMS
simulation pipeline and returns a structured report dict.

The agent:
  1. Activates the target project and study case in PowerFactory
  2. Runs a load-flow (ComLdf) to obtain initial conditions
  3. Applies the fault event sequence defined in SimulationConfig
  4. Executes the RMS simulation (ComInc + ComSim)
  5. Exports selected signals (bus voltages, generator speeds / angles) to CSV

Returns
-------
dict
    success   : bool
    csv_path  : str  — absolute path to the exported ComRes CSV
    error     : str  — human-readable error message (empty on success)
    csv_export: dict — low-level export metadata
"""

from Agent_DIgSILENT import DIgSILENTAgent, SimulationConfig


def simulation_agent(cfg: SimulationConfig) -> dict:
    print("\n" + "═" * 60)
    print("  AGENT 1 — SIMULATION AGENT")
    print("═" * 60)
    agent = DIgSILENTAgent(cfg)
    report = agent.run_pipeline()
    if report["success"]:
        print(f"[OK] Simulation done. CSV -> {report['csv_path']}")
    else:
        print("[ERROR] Simulation failed.")
    return report
