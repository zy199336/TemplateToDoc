from __future__ import annotations

from pathlib import Path
import sys


def workspace_root() -> Path:
    return Path(__file__).resolve().parents[3]


def legacy_src() -> Path:
    return workspace_root() / "TemplateToDoc" / "src"


def ensure_legacy_imports() -> None:
    path = legacy_src()
    if path.exists():
        value = str(path)
        if value not in sys.path:
            sys.path.insert(0, value)
