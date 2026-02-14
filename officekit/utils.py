from __future__ import annotations

import re
from pathlib import Path

def ensure_dir(path: str | Path) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p

_INVALID_FILENAME = re.compile(r'[\\/:*?"<>|]+')
_WINDOWS_RESERVED = {
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
}

def safe_filename(name: str, replacement: str = "_") -> str:
    """Make a string safe for Windows filenames."""
    name = name.strip()
    name = _INVALID_FILENAME.sub(replacement, name)
    name = re.sub(r"\s+", " ", name).strip(" .")
    if not name:
        name = "untitled"
    stem = name.split(".", 1)[0].upper()
    if stem in _WINDOWS_RESERVED:
        name = f"_{name}"
    return name[:180] if len(name) > 180 else name
