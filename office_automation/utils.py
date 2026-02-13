from __future__ import annotations

import re
from pathlib import Path

def ensure_dir(path: str | Path) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p

_INVALID_FILENAME = re.compile(r'[\\/:*?"<>|]+')

def safe_filename(name: str, replacement: str = "_") -> str:
    """Make a string safe for Windows filenames."""
    name = name.strip()
    name = _INVALID_FILENAME.sub(replacement, name)
    name = re.sub(r"\s+", " ", name)
    return name[:180] if len(name) > 180 else name
