"""Text / chat API providers (OpenAI-compatible)."""

from __future__ import annotations

WENWEN_BASE_URL = "https://breakout.wenwen-ai.com/v1"
XIAOQIAN_BASE_URL = "https://xiaoqian.art/v1"

TEXT_PROVIDER_PROFILES: dict[str, dict] = {
    "wenwen": {
        "label": "Wenwen（OpenAI 兼容）",
        "base_url": WENWEN_BASE_URL,
        "models": ["gpt-5-high", "gpt-5", "gpt-4o", "gpt-4o-mini"],
        "default_model": "gpt-5-high",
    },
    "xiaoqian": {
        "label": "小钱.art（OpenAI 兼容）",
        "base_url": XIAOQIAN_BASE_URL,
        "models": [
            "gpt-5.5",
            "gpt-5.4",
            "gpt-5.4-mini",
            "gpt-5.4-2026-03-05",
            "gpt-5.2",
            "gpt-5.2-pro",
            "gpt-5.2-chat-latest",
            "gpt-5.2-2025-12-11",
            "gpt-5.2-pro-2025-12-11",
            "gpt-5.3-codex",
            "codex-auto-review",
            "gpt-4o-audio-preview",
            "gpt-4o-realtime-preview",
        ],
        "default_model": "gpt-5.5",
    },
}

DEFAULT_TEXT_PROVIDER = "wenwen"


def list_text_provider_ids() -> list[str]:
    return list(TEXT_PROVIDER_PROFILES.keys())


def text_provider_label(provider_id: str) -> str:
    profile = TEXT_PROVIDER_PROFILES.get(provider_id) or TEXT_PROVIDER_PROFILES[DEFAULT_TEXT_PROVIDER]
    return str(profile.get("label") or provider_id)


def text_provider_models(provider_id: str) -> list[str]:
    profile = TEXT_PROVIDER_PROFILES.get(provider_id) or TEXT_PROVIDER_PROFILES[DEFAULT_TEXT_PROVIDER]
    models = profile.get("models")
    if isinstance(models, list) and models:
        return [str(item) for item in models]
    return [str(profile.get("default_model") or "")]


def text_provider_default_model(provider_id: str) -> str:
    profile = TEXT_PROVIDER_PROFILES.get(provider_id) or TEXT_PROVIDER_PROFILES[DEFAULT_TEXT_PROVIDER]
    return str(profile.get("default_model") or text_provider_models(provider_id)[0])


def text_provider_default_base_url(provider_id: str) -> str:
    profile = TEXT_PROVIDER_PROFILES.get(provider_id) or TEXT_PROVIDER_PROFILES[DEFAULT_TEXT_PROVIDER]
    return str(profile.get("base_url") or WENWEN_BASE_URL).rstrip("/")


def migrate_text_provider_api_keys(data: dict) -> dict[str, str]:
    keys = data.get("text_provider_api_keys")
    migrated: dict[str, str] = {}
    if isinstance(keys, dict):
        migrated = {str(k): str(v) for k, v in keys.items() if v}
    legacy = str(data.get("text_api_key") or "").strip()
    if legacy and not migrated:
        provider = str(data.get("text_provider") or DEFAULT_TEXT_PROVIDER).strip()
        if provider not in TEXT_PROVIDER_PROFILES:
            provider = DEFAULT_TEXT_PROVIDER
        migrated[provider] = legacy
    image_keys = data.get("provider_api_keys")
    if isinstance(image_keys, dict):
        xiaoqian_image_key = str(image_keys.get("xiaoqian") or "").strip()
        if xiaoqian_image_key:
            migrated.setdefault("xiaoqian", xiaoqian_image_key)
    return migrated
