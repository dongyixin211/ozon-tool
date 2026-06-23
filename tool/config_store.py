from __future__ import annotations

import json
import sys
from pathlib import Path


APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "config.json"


class ConfigStoreError(RuntimeError):
    pass


def config_exists(path: Path = CONFIG_PATH) -> bool:
    return path.exists()


def read_config_payload(path: Path = CONFIG_PATH, *, raise_on_error: bool = False) -> dict:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as exc:
        if raise_on_error:
            raise ConfigStoreError(str(exc)) from exc
        return {}
    return data if isinstance(data, dict) else {}


def write_config_payload(data: dict, path: Path = CONFIG_PATH) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
