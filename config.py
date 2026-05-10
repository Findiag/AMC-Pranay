"""
config.py — config loader.

Resolution order (later wins):
    1. Hard-coded _DEFAULTS below
    2. config.json on disk (if present)
    3. Environment variables (production should rely entirely on these)

Usage:
    from config import load_config, save_config
    cfg = load_config()
    key = cfg["anthropic_api_key"]

All callers that need an API key should call load_config() each time so
config.json edits / env-var changes are picked up without restart.
"""
import json
import os
from pathlib import Path

_BASE_DIR = Path(__file__).parent
CONFIG_PATH = _BASE_DIR / "config.json"

# DATA_DIR controls where uploads/output/chroma_db/model_cache live. On
# Railway, set DATA_DIR=/data and mount a volume there for persistence
# across deploys. Defaults to the project root for local development.
DATA_DIR = Path(os.environ.get("DATA_DIR", str(_BASE_DIR)))

# Per-subdir env overrides for finer control if needed
UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR", str(DATA_DIR / "uploads")))
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", str(DATA_DIR / "output")))
CHROMA_DIR = Path(os.environ.get("CHROMA_DIR", str(DATA_DIR / "chroma_db")))
MODEL_CACHE_DIR = Path(os.environ.get("MODEL_CACHE_DIR",
                                       str(DATA_DIR / "model_cache")))

_DEFAULTS = {
    # Claude (Anthropic) — primary LLM as of 2026-05. mapper_mode="claude"
    # routes BS/PL mapping through claude_mapper.py.
    "anthropic_api_key": "",
    "claude_model": "claude-sonnet-4-6",
    # OpenAI is kept only for backward compatibility (offline → v2_llm fallback
    # path inside bs_pl_mapper). Leave empty to disable entirely.
    "openai_api_key": "",
    "use_gpt_fallback": True,
    "balance_tolerance": 1.0,
    # When true (default), only the Standalone variant is mapped & reported.
    # Consolidated sheets are still extracted but dropped before mapping so
    # LLM cost is halved and the Report has only Standalone CY/PY columns.
    # Flip to false in config.json to get both Standalone and Consolidated.
    "standalone_only": True,
    "mapper_mode": "claude",                 # claude | v2_llm | hybrid | rules_only
    "llm_mapper_model": "gpt-4o-2024-08-06", # only used when mapper_mode=v2_llm
    "critic_model": "gpt-4o-mini",
    "enable_critic": False,
    "enable_ensemble": False,
    # Airtable upload (post-processing — uploads *_Report.xlsx files).
    # Already upserts by (company_id, end_date) so re-uploading the same FY
    # updates the existing row instead of creating duplicates.
    "airtable_api_token":         "",
    "airtable_base_id":           "",
    "airtable_bs_pl_table_id":    "",
    "airtable_user_table_id":     "",
    "airtable_company_table_id":  "",
    "airtable_user_email":        "",
    "airtable_auto_upload":       False,
    "auto_observations":          True,
    # Observations + report knobs
    # NOTE: observations.py still uses the OpenAI client — keep this an
    # OpenAI model id until that module is ported to Claude. Setting a
    # claude-* id here would cause 404s from the OpenAI endpoint.
    "observations_model":         "gpt-4o-mini",
    # Production / deployment knobs
    "api_auth_token":             "",   # require Bearer token on /api/* if set
    "cors_origins":               "*",  # comma-separated origins for /api/*
}


# Map env var → config key. Each env var, when set and non-empty, OVERRIDES
# the corresponding config.json value. Booleans accept "true"/"false"/"1"/"0";
# numbers are parsed as floats. This is the production-recommended path —
# Railway / Docker / CI inject everything via env without a config.json.
_ENV_OVERRIDES = {
    "ANTHROPIC_API_KEY":          ("anthropic_api_key",         str),
    "CLAUDE_MODEL":               ("claude_model",              str),
    "OPENAI_API_KEY":             ("openai_api_key",            str),
    "USE_GPT_FALLBACK":           ("use_gpt_fallback",          bool),
    "BALANCE_TOLERANCE":          ("balance_tolerance",         float),
    "STANDALONE_ONLY":            ("standalone_only",           bool),
    "MAPPER_MODE":                ("mapper_mode",               str),
    "LLM_MAPPER_MODEL":           ("llm_mapper_model",          str),
    "CRITIC_MODEL":               ("critic_model",              str),
    "ENABLE_CRITIC":              ("enable_critic",             bool),
    "ENABLE_ENSEMBLE":            ("enable_ensemble",           bool),
    "AIRTABLE_API_TOKEN":         ("airtable_api_token",        str),
    "AIRTABLE_BASE_ID":           ("airtable_base_id",          str),
    "AIRTABLE_BS_PL_TABLE_ID":    ("airtable_bs_pl_table_id",   str),
    "AIRTABLE_USER_TABLE_ID":     ("airtable_user_table_id",    str),
    "AIRTABLE_COMPANY_TABLE_ID":  ("airtable_company_table_id", str),
    "AIRTABLE_USER_EMAIL":        ("airtable_user_email",       str),
    "AIRTABLE_AUTO_UPLOAD":       ("airtable_auto_upload",      bool),
    "AUTO_OBSERVATIONS":          ("auto_observations",         bool),
    "OBSERVATIONS_MODEL":         ("observations_model",        str),
    "API_AUTH_TOKEN":             ("api_auth_token",            str),
    "CORS_ORIGINS":               ("cors_origins",              str),
}


def _coerce(raw: str, kind):
    """Cast a raw env-var string to the target type."""
    if kind is bool:
        return raw.strip().lower() in ("1", "true", "yes", "on")
    if kind is float:
        try: return float(raw)
        except (TypeError, ValueError): return None
    return raw.strip()


def load_config() -> dict:
    """Resolve config from defaults → config.json → env vars (last wins)."""
    cfg = dict(_DEFAULTS)

    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            # Strip _comment_ keys
            for k, v in data.items():
                if not k.startswith("_"):
                    cfg[k] = v
        except Exception:
            pass

    # Apply env-var overrides — production should rely entirely on these
    for env_name, (key, kind) in _ENV_OVERRIDES.items():
        raw = os.environ.get(env_name, "")
        if raw == "" or raw is None:
            continue
        coerced = _coerce(raw, kind)
        if coerced is not None:
            cfg[key] = coerced
    return cfg


def save_config(updates: dict) -> dict:
    """Persist updates to config.json. Keeps _comment_ keys."""
    existing = {}
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                existing = json.load(f)
        except Exception:
            existing = {}
    # Merge
    for k, v in updates.items():
        existing[k] = v
    # Write back preserving comment keys
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, indent=2)
    return load_config()


def get_api_key() -> str:
    """Shortcut: just the OpenAI API key string (legacy)."""
    return load_config().get("openai_api_key", "").strip()


def get_anthropic_key() -> str:
    """Shortcut: just the Anthropic (Claude) API key string."""
    return load_config().get("anthropic_api_key", "").strip()
