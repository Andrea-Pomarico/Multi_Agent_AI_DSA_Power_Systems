# Multi-Agent AI Pipeline for Autonomous Dynamic Security Assessment of Power Systems

This repository implements a multi-agent AI pipeline for autonomous power system stability analysis in **DIgSILENT PowerFactory**.

It covers the full workflow—from fault simulation to polished engineering reports—leveraging LLM agents (e.g., Google Gemini) for narrative generation, validation, and mitigation planning.

```
simulation_config.json
        │
   [Agent 0] Intake          — natural language → structured cases
        │
   [Agent 1] Simulation      — PowerFactory RMS → CSV export
        │
   [Agent 2] Analysis        — voltages / speeds / angles → KPIs
        ├─────────────────────────────────────┐
   [Agent 3] Plot            — PNG time-series [Agent 4] Summary      — LLM draft
                                               [Agent 5] Review       — LLM fact-check
                                               [Agent 6] Final report — LLM polish
        └─────────────────────────────────────┘
   [Agent 7] Comparison      — cross-case risk assessment
        ├──────────────────┬──────────────────┐
   [Agent 8] Presentation  [Agent 9] Mitigation  report_utils
   (.pptx)                 (ranked actions)      (.docx / CSV / TXT)
```

---

## Features

- **Multi-case automation** — define any number of fault scenarios in a single JSON config; the pipeline runs them sequentially and produces a unified comparative report
- **Natural-language intake** — describe scenarios in plain text; Agent 0 parses them into structured cases via LLM
- **LLM-powered reporting** with a draft → review → final loop that cross-checks every numerical claim against computed KPIs before publishing
- **Zero python-docx dependency** — Word reports are generated with stdlib `zipfile` + raw OOXML
- **Automatic plots** — voltage magnitude, generator speed, and per-bus violation charts; multi-case overlay comparisons

---

## Project structure

```
.
├── Multi_Agent_RMS_google.py      # Main entry point
├── Agent_DIgSILENT.py             # PowerFactory wrapper + SimulationConfig
├── agent_intake.py                # Agent 0
├── agent_simulation.py            # Agent 1
├── agent_analysis.py              # Agent 2
├── agent_plot.py                  # Agent 3
├── agent_llm_reporting.py         # Agents 4–7
├── agent_presentation.py          # Agent 8
├── agent_mitigation.py            # Agent 9
├── llm_client.py                  # Provider-agnostic LLM wrapper
├── prompt_loader.py               # Loads system prompts from CSV
├── report_utils.py                # Disk persistence helpers
├── agent_prompts.csv              # All LLM system prompts (editable)
├── simulation_config_example.json # Template config (safe to commit)
├── simulation_config.json         # Your local config (gitignored)
├── requirements.txt
└── README.md
```

---
