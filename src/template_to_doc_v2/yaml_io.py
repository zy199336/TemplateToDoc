from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml


class ConfigError(RuntimeError):
    pass


def load_yaml(path: str | Path) -> dict[str, Any]:
    target = Path(path)
    if not target.exists():
        raise ConfigError(f"YAML file not found: {target}")
    data = yaml.safe_load(target.read_text(encoding="utf-8")) or {}
    if not isinstance(data, dict):
        raise ConfigError(f"YAML file must contain a mapping: {target}")
    return data


def write_yaml(path: str | Path, data: dict[str, Any]) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding="utf-8")
    return target


def dotted_get(data: dict[str, Any], path: str, default: Any = "") -> Any:
    current: Any = data
    for part in str(path or "").split("."):
        if not isinstance(current, dict) or part not in current:
            return default
        current = current[part]
    return default if current in (None, "") else current


def dotted_set(data: dict[str, Any], path: str, value: Any) -> None:
    current = data
    parts = str(path or "").split(".")
    for part in parts[:-1]:
        child = current.setdefault(part, {})
        if not isinstance(child, dict):
            raise ConfigError(f"Cannot set nested field under non-mapping: {path}")
        current = child
    current[parts[-1]] = value
