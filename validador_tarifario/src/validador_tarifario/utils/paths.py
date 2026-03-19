from __future__ import annotations

from pathlib import Path

PROJECT_NAME = "Validador Tarifario"
SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
DEFAULT_SHEET_NAME = "AGRUPADOR"
ERROR_SHEET_NAME = "ERRORES"
DEFAULT_SCAN_ROWS = 15
REFERENCE_DATA_DIR = Path(__file__).resolve().parents[4] / "base de datos"
SOAT_UVB_2026_FILE = REFERENCE_DATA_DIR / "SOAT UVB 2026.xlsx"
SOAT_2025_FILE = REFERENCE_DATA_DIR / "SOAT 2025.xlsx"


def ensure_output_dir(path: str | Path) -> Path:
    output_dir = Path(path).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir
