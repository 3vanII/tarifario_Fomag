from __future__ import annotations

from pathlib import Path

PROJECT_NAME = "Validador Tarifario"
SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
DEFAULT_SHEET_NAME = "AGRUPADOR"
DEFAULT_SCAN_ROWS = 15


def ensure_output_dir(path: str | Path) -> Path:
    output_dir = Path(path).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir
