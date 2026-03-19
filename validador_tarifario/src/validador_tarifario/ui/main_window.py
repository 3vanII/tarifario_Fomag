from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from validador_tarifario.services.agrupador_service import (
    AgrupadorConfig,
    AgrupadorError,
    ExcelAgrupadorService,
)
from validador_tarifario.utils.paths import DEFAULT_SHEET_NAME, PROJECT_NAME


class MainWindow(ttk.Frame):
    def __init__(self, master: tk.Tk) -> None:
        super().__init__(master, padding=16)
        self.master = master
        self.pack(fill="both", expand=True)

        self.input_file_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.sheet_name_var = tk.StringVar(value=DEFAULT_SHEET_NAME)
        self.status_var = tk.StringVar(value="Listo para iniciar.")

        self._build_styles()
        self._build_ui()

    def _build_styles(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Hint.TLabel", foreground="#4b5563")
        style.configure("Primary.TButton", padding=(10, 8))

    def _build_ui(self) -> None:
        self.columnconfigure(1, weight=1)
        self.rowconfigure(5, weight=1)

        ttk.Label(self, text=PROJECT_NAME, style="Title.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w"
        )
        ttk.Label(
            self,
            text="Agrupa todas las hojas del Excel en una nueva hoja y agrega la columna 'Pestaña origen'.",
            style="Hint.TLabel",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 16))

        ttk.Label(self, text="Archivo Excel:").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(self, textvariable=self.input_file_var).grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Button(self, text="Examinar", command=self.select_input_file).grid(row=2, column=2, sticky="ew")

        ttk.Label(self, text="Carpeta salida:").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Entry(self, textvariable=self.output_dir_var).grid(row=3, column=1, sticky="ew", padx=8)
        ttk.Button(self, text="Elegir", command=self.select_output_dir).grid(row=3, column=2, sticky="ew")

        ttk.Label(self, text="Nombre hoja agrupadora:").grid(row=4, column=0, sticky="w", pady=6)
        ttk.Entry(self, textvariable=self.sheet_name_var).grid(row=4, column=1, sticky="ew", padx=8)
        ttk.Button(self, text="Usar AGRUPADOR", command=self.restore_default_sheet_name).grid(
            row=4, column=2, sticky="ew"
        )

        log_frame = ttk.LabelFrame(self, text="Log de ejecución", padding=8)
        log_frame.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(16, 12))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, wrap="word", height=16, state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        bottom_frame = ttk.Frame(self)
        bottom_frame.grid(row=6, column=0, columnspan=3, sticky="ew")
        bottom_frame.columnconfigure(0, weight=1)

        ttk.Label(bottom_frame, textvariable=self.status_var).grid(row=0, column=0, sticky="w")
        self.process_button = ttk.Button(
            bottom_frame,
            text="Procesar archivo",
            style="Primary.TButton",
            command=self.process_file,
        )
        self.process_button.grid(row=0, column=1, sticky="e")

    def restore_default_sheet_name(self) -> None:
        self.sheet_name_var.set(DEFAULT_SHEET_NAME)

    def select_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xlsm")],
        )
        if file_path:
            self.input_file_var.set(file_path)
            if not self.output_dir_var.get().strip():
                self.output_dir_var.set(str(Path(file_path).parent / "salida"))

    def select_output_dir(self) -> None:
        folder_path = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if folder_path:
            self.output_dir_var.set(folder_path)

    def append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.master.update_idletasks()

    def set_processing_state(self, processing: bool) -> None:
        state = "disabled" if processing else "normal"
        self.process_button.configure(state=state)

    def validate_form(self) -> bool:
        input_file = self.input_file_var.get().strip()
        output_dir = self.output_dir_var.get().strip()
        sheet_name = self.sheet_name_var.get().strip()

        if not input_file:
            messagebox.showwarning(PROJECT_NAME, "Debes seleccionar un archivo Excel.")
            return False
        if not output_dir:
            messagebox.showwarning(PROJECT_NAME, "Debes seleccionar una carpeta de salida.")
            return False
        if not sheet_name:
            messagebox.showwarning(PROJECT_NAME, "Debes indicar el nombre de la hoja agrupadora.")
            return False
        return True

    def process_file(self) -> None:
        if not self.validate_form():
            return

        self.set_processing_state(True)
        self.status_var.set("Procesando archivo...")
        self.append_log("=" * 70)
        self.append_log("Inicio del proceso")

        worker = threading.Thread(target=self._run_process, daemon=True)
        worker.start()

    def _run_process(self) -> None:
        try:
            config = AgrupadorConfig(
                input_file=Path(self.input_file_var.get().strip()),
                output_dir=Path(self.output_dir_var.get().strip()),
                sheet_name=self.sheet_name_var.get().strip(),
            )
            service = ExcelAgrupadorService(logger=self._thread_safe_log)
            result = service.process_file(config)

            self.master.after(0, self._on_success, result.output_file, result.total_rows, result.processed_sheets)
        except AgrupadorError as exc:
            self.master.after(0, self._on_error, str(exc))
        except Exception as exc:  # noqa: BLE001
            self.master.after(0, self._on_error, f"Error inesperado: {exc}")

    def _thread_safe_log(self, message: str) -> None:
        self.master.after(0, self.append_log, message)

    def _on_success(self, output_file: Path, total_rows: int, processed_sheets: int) -> None:
        self.append_log(f"Hojas procesadas: {processed_sheets}")
        self.append_log(f"Total filas agrupadas: {total_rows:,}")
        self.append_log(f"Archivo de salida: {output_file}")
        self.append_log("Proceso finalizado correctamente.")
        self.status_var.set("Proceso completado.")
        self.set_processing_state(False)
        messagebox.showinfo(
            PROJECT_NAME,
            f"Proceso completado correctamente.\n\nArchivo generado:\n{output_file}",
        )

    def _on_error(self, message: str) -> None:
        self.append_log(message)
        self.status_var.set("Se produjo un error.")
        self.set_processing_state(False)
        messagebox.showerror(PROJECT_NAME, message)
