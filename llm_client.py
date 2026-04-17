"""
LLM Client
==========
Provider-agnostic wrapper that routes agent calls to either Google Gemini
or Groq, based on the active configuration.

Supported providers
-------------------
  gemini  — Google Generative AI SDK (google-genai)
  groq    — Groq Python SDK

Model aliases
-------------
  MODEL_FAST  ("fast")   → gemini-2.5-flash-lite  /  llama-3.1-8b-instant
  MODEL_SMART ("smart")  → gemini-2.5-flash        /  llama-3.3-70b-versatile

Configuration
-------------
  Call configure_llm_from_config(raw_cfg) once at startup with the parsed
  simulation_config.json dict.  API keys and model names can also be
  supplied via environment variables:
      GEMINI_API_KEY, GROQ_API_KEY,
      GEMINI_MODEL_FAST, GEMINI_MODEL_SMART,
      GROQ_MODEL_FAST,  GROQ_MODEL_SMART

Public API
----------
  configure_llm_from_config(raw_cfg: dict) → None
  get_llm_runtime_info()                   → dict
  run_agent(system_prompt, user_message, max_tokens, model) → str
"""

import os

from google import genai
from google.genai import types
from groq import Groq

# Logical model aliases used by agents.
MODEL_FAST  = "fast"
MODEL_SMART = "smart"

_state = {
    "provider": os.getenv("LLM_PROVIDER", "gemini").strip().lower(),
    "gemini_api_key": os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY", ""),
    "groq_api_key": os.getenv("GROQ_API_KEY", ""),
    "gemini_models": {
        "fast": os.getenv("GEMINI_MODEL_FAST", "gemini-2.5-flash-lite"),
        "smart": os.getenv("GEMINI_MODEL_SMART", "gemini-2.5-flash"),
    },
    "groq_models": {
        "fast": os.getenv("GROQ_MODEL_FAST", "llama-3.1-8b-instant"),
        "smart": os.getenv("GROQ_MODEL_SMART", "llama-3.3-70b-versatile"),
    },
}

_gemini_client = None
_groq_client = None


def configure_llm_from_config(raw_cfg: dict | None) -> None:
    """
    Configure LLM provider, keys, and model names from simulation_config JSON.

        Supported structures:
    {
      "llm": {
        "provider": "gemini" | "groq",
                "api_keys": {
                    "gemini": "...",
                    "groq": "..."
                },
                "models_by_provider": {
                    "gemini": {"fast": "...", "smart": "..."},
                    "groq": {"fast": "...", "smart": "..."}
                }
      }
    }
        Backward-compatible keys are still accepted:
            - gemini_api_key / groq_api_key
            - models (active provider only)
            - gemini_models / groq_models
    """
    global _gemini_client, _groq_client

    if not isinstance(raw_cfg, dict):
        return

    llm_cfg = raw_cfg.get("llm", raw_cfg)
    if not isinstance(llm_cfg, dict):
        return

    provider = str(llm_cfg.get("provider", _state["provider"]))
    provider = provider.strip().lower()
    if provider in ("gemini", "groq"):
        _state["provider"] = provider

    # New style: provider-key map.
    api_keys = llm_cfg.get("api_keys")
    if isinstance(api_keys, dict):
        if isinstance(api_keys.get("gemini"), str) and api_keys["gemini"].strip():
            _state["gemini_api_key"] = api_keys["gemini"].strip()
        if isinstance(api_keys.get("groq"), str) and api_keys["groq"].strip():
            _state["groq_api_key"] = api_keys["groq"].strip()

    # Backward-compatible key fields.
    gemini_key = llm_cfg.get("gemini_api_key") or llm_cfg.get("google_api_key")
    if isinstance(gemini_key, str) and gemini_key.strip():
        _state["gemini_api_key"] = gemini_key.strip()

    groq_key = llm_cfg.get("groq_api_key")
    if isinstance(groq_key, str) and groq_key.strip():
        _state["groq_api_key"] = groq_key.strip()

    # New style: provider-model map.
    models_by_provider = llm_cfg.get("models_by_provider")
    if isinstance(models_by_provider, dict):
        gemini_models = models_by_provider.get("gemini")
        if isinstance(gemini_models, dict):
            if isinstance(gemini_models.get("fast"), str) and gemini_models["fast"].strip():
                _state["gemini_models"]["fast"] = gemini_models["fast"].strip()
            if isinstance(gemini_models.get("smart"), str) and gemini_models["smart"].strip():
                _state["gemini_models"]["smart"] = gemini_models["smart"].strip()

        groq_models = models_by_provider.get("groq")
        if isinstance(groq_models, dict):
            if isinstance(groq_models.get("fast"), str) and groq_models["fast"].strip():
                _state["groq_models"]["fast"] = groq_models["fast"].strip()
            if isinstance(groq_models.get("smart"), str) and groq_models["smart"].strip():
                _state["groq_models"]["smart"] = groq_models["smart"].strip()

    # Backward-compatible generic model block applies only to active provider.
    models_cfg = llm_cfg.get("models")
    if isinstance(models_cfg, dict):
        fast = models_cfg.get("fast")
        smart = models_cfg.get("smart")
        if isinstance(fast, str) and fast.strip():
            if _state["provider"] == "gemini":
                _state["gemini_models"]["fast"] = fast.strip()
            else:
                _state["groq_models"]["fast"] = fast.strip()
        if isinstance(smart, str) and smart.strip():
            if _state["provider"] == "gemini":
                _state["gemini_models"]["smart"] = smart.strip()
            else:
                _state["groq_models"]["smart"] = smart.strip()

    # Backward-compatible provider-specific model blocks.
    gm = llm_cfg.get("gemini_models")
    if isinstance(gm, dict):
        if isinstance(gm.get("fast"), str) and gm["fast"].strip():
            _state["gemini_models"]["fast"] = gm["fast"].strip()
        if isinstance(gm.get("smart"), str) and gm["smart"].strip():
            _state["gemini_models"]["smart"] = gm["smart"].strip()

    grm = llm_cfg.get("groq_models")
    if isinstance(grm, dict):
        if isinstance(grm.get("fast"), str) and grm["fast"].strip():
            _state["groq_models"]["fast"] = grm["fast"].strip()
        if isinstance(grm.get("smart"), str) and grm["smart"].strip():
            _state["groq_models"]["smart"] = grm["smart"].strip()

    # Invalidate clients in case key/provider changed.
    _gemini_client = None
    _groq_client = None


def get_llm_runtime_info() -> dict:
    provider = _state["provider"]
    model_fast = _state["gemini_models"]["fast"] if provider == "gemini" else _state["groq_models"]["fast"]
    model_smart = _state["gemini_models"]["smart"] if provider == "gemini" else _state["groq_models"]["smart"]
    return {
        "provider": provider,
        "model_fast": model_fast,
        "model_smart": model_smart,
    }


def _resolve_model(model: str) -> str:
    provider = _state["provider"]
    if model in (MODEL_FAST, "fast"):
        return _state["gemini_models"]["fast"] if provider == "gemini" else _state["groq_models"]["fast"]
    if model in (MODEL_SMART, "smart"):
        return _state["gemini_models"]["smart"] if provider == "gemini" else _state["groq_models"]["smart"]
    return model


def _get_gemini_client():
    global _gemini_client
    if _gemini_client is not None:
        return _gemini_client

    api_key = _state["gemini_api_key"]
    if not api_key:
        raise RuntimeError(
            "Gemini selected but no API key configured. "
            "Set llm.gemini_api_key in simulation_config.json or GEMINI_API_KEY env var."
        )

    _gemini_client = genai.Client(api_key=api_key)
    return _gemini_client


def _get_groq_client():
    global _groq_client
    if _groq_client is not None:
        return _groq_client

    api_key = _state["groq_api_key"]
    if not api_key:
        raise RuntimeError(
            "Groq selected but no API key configured. "
            "Set llm.groq_api_key in simulation_config.json or GROQ_API_KEY env var."
        )

    _groq_client = Groq(api_key=api_key)
    return _groq_client


def run_vision_agent(system_prompt: str, user_message: str,
                     image_path: str | None = None,
                     max_tokens: int = 2000,
                     model: str = MODEL_SMART) -> str:
    """
    Like run_agent, but optionally attaches a PNG image to the request.

    Image support per provider
    --------------------------
    Gemini  : full multimodal — image is embedded in the content parts.
    Groq    : text-only fallback — the image is silently ignored and only
              the text message is sent (Groq's default llama models do not
              support vision; swap to a vision-capable model if needed).

    Parameters
    ----------
    system_prompt : str
    user_message  : str
    image_path    : str or None — absolute path to a PNG/JPG to include.
    max_tokens    : int
    model         : str — MODEL_FAST or MODEL_SMART alias, or a literal name.

    Returns
    -------
    str — LLM response text.
    """
    import base64

    provider       = _state["provider"]
    resolved_model = _resolve_model(model)

    # ── Gemini: multimodal request ────────────────────────────────
    if provider == "gemini":
        contents = []
        if image_path and os.path.isfile(image_path):
            with open(image_path, "rb") as f:
                img_bytes = f.read()
            mime = "image/png" if image_path.lower().endswith(".png") else "image/jpeg"
            contents.append(types.Part.from_bytes(data=img_bytes, mime_type=mime))
        contents.append(types.Part.from_text(text=user_message))

        response = _get_gemini_client().models.generate_content(
            model=resolved_model,
            contents=contents,
            config=types.GenerateContentConfig(
                system_instruction=system_prompt,
                max_output_tokens=max_tokens,
                temperature=0.1,
            ),
        )
        return response.text or ""

    # ── Groq: text-only fallback (vision not supported on default models) ──
    return _get_groq_client().chat.completions.create(
        model=resolved_model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user",   "content": user_message},
        ],
        max_tokens=max_tokens,
        temperature=0.1,
    ).choices[0].message.content


def run_agent(system_prompt: str, user_message: str,
              max_tokens: int = 700, model: str = MODEL_FAST) -> str:
    provider = _state["provider"]
    resolved_model = _resolve_model(model)

    if provider == "gemini":
        response = _get_gemini_client().models.generate_content(
            model=resolved_model,
            contents=user_message,
            config=types.GenerateContentConfig(
                system_instruction=system_prompt,
                max_output_tokens=max_tokens,
                temperature=0.1,
            ),
        )
        return response.text or ""

    response = _get_groq_client().chat.completions.create(
        model=resolved_model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message},
        ],
        max_tokens=max_tokens,
        temperature=0.1,
    )
    return response.choices[0].message.content
