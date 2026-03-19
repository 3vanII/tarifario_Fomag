from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Callable

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from validador_tarifario.utils.paths import (
    DEFAULT_SCAN_ROWS,
    ERROR_SHEET_NAME,
    SOAT_2025_FILE,
    SOAT_UVB_2026_FILE,
    SUPPORTED_EXTENSIONS,
    TABLA_REFERENCIA_CUPS_FILE,
)

LogFn = Callable[[str], None]


class AgrupadorError(Exception):
    pass


@dataclass(slots=True)
class AgrupadorConfig:
    input_file: Path
    output_dir: Path
    sheet_name: str = "AGRUPADOR"
    scan_rows: int = DEFAULT_SCAN_ROWS


@dataclass(slots=True)
class AgrupadorResult:
    output_file: Path
    total_sheets: int
    processed_sheets: int
    total_rows: int
    error_rows: int
    total_input_rows: int


class ExcelAgrupadorService:
    def __init__(self, logger: LogFn | None = None) -> None:
        self.logger = logger or (lambda msg: None)
        self._reference_codes_cache: set[str] | None = None

    def log(self, message: str) -> None:
        self.logger(message)

    def process_file(self, config: AgrupadorConfig) -> AgrupadorResult:
        input_file = config.input_file.expanduser().resolve()
        output_dir = config.output_dir.expanduser().resolve()

        self._validate_input(input_file)
        output_dir.mkdir(parents=True, exist_ok=True)

        self.log(f"Leyendo libro: {input_file.name}")
        xls = pd.ExcelFile(input_file)
        all_sheets = list(xls.sheet_names)
        target_sheets = [s for s in all_sheets if s.strip().lower() != config.sheet_name.strip().lower()]

        if not target_sheets:
            raise AgrupadorError("No hay hojas disponibles para agrupar.")

        dataframes: list[pd.DataFrame] = []

        for sheet_name in target_sheets:
            self.log(f"Procesando hoja: {sheet_name}")
            try:
                df = self._read_normalized_sheet(
                    file_path=input_file,
                    sheet_name=sheet_name,
                    scan_rows=config.scan_rows,
                )
                if df.empty:
                    self.log(f"  - Hoja omitida por no contener datos útiles: {sheet_name}")
                    continue

                df.insert(0, "Pestaña origen", sheet_name)
                dataframes.append(df)
                self.log(f"  - Filas agregadas: {len(df):,}")
            except Exception as exc:  # noqa: BLE001
                self.log(f"  - Error en hoja '{sheet_name}': {exc}")

        if not dataframes:
            raise AgrupadorError("No se encontraron datos útiles en ninguna hoja del archivo.")

        grouped_df = pd.concat(dataframes, ignore_index=True, sort=False)
        valid_df, error_df = self._validate_cups_codes(grouped_df)

        output_file = output_dir / input_file.name
        self._write_grouped_sheet(
            input_file=input_file,
            output_file=output_file,
            grouped_df=valid_df,
            error_df=error_df,
            grouped_sheet_name=config.sheet_name,
        )

        self.log(f"Archivo generado correctamente: {output_file}")
        self.log(
            f"Filas válidas en '{config.sheet_name}': {len(valid_df):,} • "
            f"Filas con error en '{ERROR_SHEET_NAME}': {len(error_df):,}"
        )
        return AgrupadorResult(
            output_file=output_file,
            total_sheets=len(all_sheets),
            processed_sheets=len(dataframes),
            total_rows=len(valid_df),
            error_rows=len(error_df),
            total_input_rows=len(grouped_df),
        )

    def _validate_input(self, input_file: Path) -> None:
        if not input_file.exists():
            raise AgrupadorError(f"El archivo no existe: {input_file}")
        if not input_file.is_file():
            raise AgrupadorError("La ruta seleccionada no es un archivo.")
        if input_file.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise AgrupadorError("Solo se permiten archivos .xlsx o .xlsm")

    @staticmethod
    def _is_empty(value) -> bool:
        if pd.isna(value):
            return True
        if isinstance(value, str) and value.strip() == "":
            return True
        return False

    @staticmethod
    def _clean_text(value) -> str:
        if value is None or pd.isna(value):
            return ""
        return str(value).strip()

    @staticmethod
    def _normalize_code(value) -> str:
        if value is None or pd.isna(value):
            return ""

        if isinstance(value, int):
            return str(value)

        if isinstance(value, float):
            if value.is_integer():
                return str(int(value))
            return str(value).strip()

        text = str(value).strip()
        if not text:
            return ""

        if re.fullmatch(r"\d+\.0+", text):
            return text.split(".", maxsplit=1)[0]

        return text

    def _make_unique_columns(self, columns) -> list[str]:
        unique_names: list[str] = []
        counts: dict[str, int] = {}

        for index, column in enumerate(columns, start=1):
            name = self._clean_text(column)
            if not name:
                name = f"COL_{index}"

            name = re.sub(r"\s+", " ", name).strip()

            if name in counts:
                counts[name] += 1
                final_name = f"{name}_{counts[name]}"
            else:
                counts[name] = 1
                final_name = name

            unique_names.append(final_name)

        return unique_names

    def _detect_header_row(self, file_path: Path, sheet_name: str, scan_rows: int) -> int:
        preview = pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            header=None,
            nrows=scan_rows,
            dtype=object,
        )

        best_row = 0
        best_score = -1

        for index, row in preview.iterrows():
            values = [v for v in row.tolist() if not self._is_empty(v)]
            if not values:
                continue

            non_empty_count = len(values)
            text_count = sum(1 for v in values if isinstance(v, str))
            unique_count = len({str(v).strip().lower() for v in values})

            score = (non_empty_count * 10) + (text_count * 3) + unique_count
            if score > best_score:
                best_score = score
                best_row = int(index)

        return best_row

    def _read_normalized_sheet(self, file_path: Path, sheet_name: str, scan_rows: int) -> pd.DataFrame:
        header_row = self._detect_header_row(file_path, sheet_name, scan_rows)

        df = pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            header=header_row,
            dtype=object,
        )

        df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")

        if df.empty:
            return pd.DataFrame()

        df.columns = self._make_unique_columns(df.columns)
        df = df.dropna(axis=0, how="all")
        return df

    def _load_reference_file(self, file_path: Path) -> pd.DataFrame:
        suffix = file_path.suffix.lower()
        if suffix == ".csv":
            return pd.read_csv(file_path, dtype=object, sep=None, engine="python")
        return pd.read_excel(file_path, dtype=object)

    def _get_reference_codes(self) -> set[str]:
        if self._reference_codes_cache is not None:
            return self._reference_codes_cache

        reference_sources = [
            (SOAT_UVB_2026_FILE, "CODIGO"),
            (SOAT_2025_FILE, "CÓDIGO"),
            (TABLA_REFERENCIA_CUPS_FILE, "Codigo"),
        ]

        reference_codes: set[str] = set()

        for file_path, column_name in reference_sources:
            if not file_path.exists():
                raise AgrupadorError(f"No se encontró la base de referencia: {file_path}")

            self.log(f"Cargando base de referencia: {file_path.name}")
            df = self._load_reference_file(file_path)

            matched_column = self._find_column_name(df, column_name)
            if matched_column is None:
                raise AgrupadorError(
                    f"La columna '{column_name}' no existe en la base de referencia {file_path.name}."
                )

            reference_codes.update(
                normalized_code
                for normalized_code in (self._normalize_code(value) for value in df[matched_column].tolist())
                if normalized_code
            )

        self._reference_codes_cache = reference_codes
        return reference_codes

    def _find_column_name(self, df: pd.DataFrame, expected_name: str) -> str | None:
        expected = expected_name.strip().upper()
        for column in df.columns:
            if self._clean_text(column).upper() == expected:
                return column
        return None

    def _validate_cups_codes(self, grouped_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
        cups_column = self._find_column_name(grouped_df, "CUPS")
        if cups_column is None:
            raise AgrupadorError("No se encontró la columna 'CUPS' en el archivo a procesar.")

        reference_codes = self._get_reference_codes()
        normalized_codes = grouped_df[cups_column].map(self._normalize_code)
        invalid_mask = ~normalized_codes.isin(reference_codes)

        error_df = grouped_df.loc[invalid_mask].copy()
        if not error_df.empty:
            error_df["Motivo del error"] = "Codigo CUPS invalido"
            self.log(f"Validación CUPS: {len(error_df):,} fila(s) enviadas a la hoja de errores.")
        else:
            self.log("Validación CUPS: no se encontraron códigos inválidos.")

        valid_df = grouped_df.loc[~invalid_mask].copy()
        return valid_df, error_df

    @staticmethod
    def _clean_excel_value(value):
        if pd.isna(value):
            return None
        return value

    def _append_dataframe_to_worksheet(self, worksheet, dataframe: pd.DataFrame) -> None:
        for row in dataframe_to_rows(dataframe, index=False, header=True):
            worksheet.append([self._clean_excel_value(value) for value in row])

        if worksheet.max_row >= 1 and worksheet.max_column >= 1:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                cell_value = "" if cell.value is None else str(cell.value)
                if len(cell_value) > max_length:
                    max_length = len(cell_value)
            worksheet.column_dimensions[column_letter].width = min(max_length + 2, 40)

    def _write_grouped_sheet(
        self,
        input_file: Path,
        output_file: Path,
        grouped_df: pd.DataFrame,
        error_df: pd.DataFrame,
        grouped_sheet_name: str,
    ) -> None:
        keep_vba = output_file.suffix.lower() == ".xlsm"
        workbook = load_workbook(input_file, keep_vba=keep_vba)

        target_name = grouped_sheet_name.strip() or "AGRUPADOR"

        for sheet_name in list(workbook.sheetnames):
            workbook.remove(workbook[sheet_name])

        worksheet = workbook.create_sheet(title=target_name)
        self._append_dataframe_to_worksheet(worksheet, grouped_df)

        if not error_df.empty:
            error_sheet_name = ERROR_SHEET_NAME
            if error_sheet_name == target_name:
                error_sheet_name = f"{ERROR_SHEET_NAME}_1"
            error_worksheet = workbook.create_sheet(title=error_sheet_name)
            self._append_dataframe_to_worksheet(error_worksheet, error_df)

        workbook.save(output_file)
