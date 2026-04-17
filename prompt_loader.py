"""
Prompt Loader — Load system prompts from CSV file for easy modification.
"""

import os
import csv
from typing import Dict, Optional


_prompt_cache: Dict[str, str] = {}


def load_prompts(csv_path: str | None = None) -> Dict[str, str]:
    """
    Load all agent system prompts from CSV file.
    
    CSV format:
        agent_id, agent_name, system_prompt
    
    Args:
        csv_path: Path to agent_prompts.csv. If None, uses default location.
    
    Returns:
        Dict mapping agent_id -> system_prompt
    """
    global _prompt_cache
    
    if csv_path is None:
        csv_path = os.path.join(os.path.dirname(__file__), "agent_prompts.csv")
    
    if not os.path.isfile(csv_path):
        print(f"[WARN] Prompt CSV not found: {csv_path}")
        return {}
    
    prompts: Dict[str, str] = {}
    try:
        with open(csv_path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row and "agent_id" in row and "system_prompt" in row:
                    agent_id = row["agent_id"].strip()
                    prompt = row["system_prompt"].strip()
                    if agent_id and prompt:
                        prompts[agent_id] = prompt
        
        _prompt_cache = prompts
        print(f"[OK] Loaded {len(prompts)} prompts from {csv_path}")
        return prompts
    except Exception as e:
        print(f"[ERROR] Failed to load prompts from {csv_path}: {e}")
        return {}


def get_prompt(agent_id: str, default: str = "") -> str:
    """
    Get system prompt for a given agent ID.
    
    Args:
        agent_id: e.g., "agent_4_summary", "agent_7_comparison"
        default: Default prompt if agent_id not found
    
    Returns:
        System prompt string
    """
    if not _prompt_cache:
        load_prompts()
    
    return _prompt_cache.get(agent_id, default)
