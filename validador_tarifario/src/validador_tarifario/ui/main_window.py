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
        super().__init__(master, padding=0, style="App.TFrame")
        self.master = master
        self.pack(fill="both", expand=True)

        self.input_file_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.sheet_name_var = tk.StringVar(value=DEFAULT_SHEET_NAME)
        self.status_var = tk.StringVar(value="Listo para iniciar.")
        self.input_summary_var = tk.StringVar(value="Aún no has seleccionado un archivo.")
        self.output_summary_var = tk.StringVar(value="Se creará una carpeta de salida cuando elijas un archivo.")
        self.sheet_summary_var = tk.StringVar(value=f"El Excel generado tendrá una sola hoja: {DEFAULT_SHEET_NAME}.")

        self._build_styles()
        self._build_ui()
        self.sheet_name_var.trace_add("write", lambda *_: self._refresh_summary())
        self._refresh_summary()

    def _build_styles(self) -> None:
        self.master.configure(bg="#eef4ff")

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("App.TFrame", background="#eef4ff")
        style.configure("Card.TFrame", background="#ffffff", relief="flat")
        style.configure("Section.TLabelframe", background="#ffffff", borderwidth=0)
        style.configure("Section.TLabelframe.Label", background="#ffffff", foreground="#0f172a", font=("Segoe UI", 11, "bold"))
        style.configure("Title.TLabel", background="#eef4ff", foreground="#0f172a", font=("Segoe UI", 22, "bold"))
        style.configure("Hero.TLabel", background="#eef4ff", foreground="#475569", font=("Segoe UI", 11))
        style.configure("MetricValue.TLabel", background="#ffffff", foreground="#1d4ed8", font=("Segoe UI", 18, "bold"))
        style.configure("MetricLabel.TLabel", background="#ffffff", foreground="#64748b", font=("Segoe UI", 9))
        style.configure("FieldLabel.TLabel", background="#ffffff", foreground="#334155", font=("Segoe UI", 10, "bold"))
        style.configure("Hint.TLabel", background="#ffffff", foreground="#64748b", font=("Segoe UI", 9))
        style.configure("Status.TLabel", background="#eef4ff", foreground="#0f172a", font=("Segoe UI", 10, "bold"))
        style.configure("Primary.TButton", padding=(14, 10), font=("Segoe UI", 10, "bold"))
        style.configure("Secondary.TButton", padding=(12, 8), font=("Segoe UI", 9, "bold"))
        style.configure("Modern.Horizontal.TProgressbar", troughcolor="#dbeafe", background="#2563eb", bordercolor="#dbeafe", lightcolor="#2563eb", darkcolor="#2563eb")
        style.map(
            "Primary.TButton",
            background=[("active", "#1d4ed8"), ("!disabled", "#2563eb")],
            foreground=[("!disabled", "white")],
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#dbeafe"), ("!disabled", "#eff6ff")],
            foreground=[("!disabled", "#1d4ed8")],
        )

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        self._build_header()
        self._build_metrics()
        self._build_content()
        self._build_footer()

    def _build_header(self) -> None:
        header = ttk.Frame(self, style="App.TFrame", padding=(22, 20, 22, 8))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text=PROJECT_NAME, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text=(
                "Convierte un libro con múltiples pestañas en un archivo final limpio, "
                "con una sola hoja agrupadora y un flujo más claro para el usuario."
            ),
            style="Hero.TLabel",
            wraplength=900,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

    def _build_metrics(self) -> None:
        metrics = ttk.Frame(self, style="App.TFrame", padding=(22, 6, 22, 12))
        metrics.grid(row=1, column=0, sticky="ew")
        for column in range(3):
            metrics.columnconfigure(column, weight=1)

        cards = [
            ("1", "Archivo de entrada", "Selecciona un Excel .xlsx o .xlsm"),
            ("1", "Hoja de salida", "Se genera únicamente la hoja solicitada"),
            ("+1", "Columna extra", "Siempre agrega 'Pestaña origen'"),
        ]

        for index, (value, title, description) in enumerate(cards):
            card = ttk.Frame(metrics, style="Card.TFrame", padding=16)
            card.grid(row=0, column=index, sticky="nsew", padx=(0 if index == 0 else 6, 0))
            card.columnconfigure(0, weight=1)
            ttk.Label(card, text=value, style="MetricValue.TLabel").grid(row=0, column=0, sticky="w")
            ttk.Label(card, text=title, style="FieldLabel.TLabel").grid(row=1, column=0, sticky="w", pady=(6, 0))
            ttk.Label(card, text=description, style="Hint.TLabel", wraplength=220, justify="left").grid(
                row=2, column=0, sticky="w", pady=(4, 0)
            )

    def _build_content(self) -> None:
        content = ttk.Frame(self, style="App.TFrame", padding=(22, 0, 22, 12))
        content.grid(row=2, column=0, sticky="nsew")
        content.columnconfigure(0, weight=5)
        content.columnconfigure(1, weight=4)
        content.rowconfigure(0, weight=1)

        form_card = ttk.Frame(content, style="Card.TFrame", padding=20)
        form_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        form_card.columnconfigure(1, weight=1)
        form_card.rowconfigure(8, weight=1)

        ttk.Label(form_card, text="Configuración del proceso", style="FieldLabel.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w"
        )
        ttk.Label(
            form_card,
            text="Define el archivo a leer, la carpeta destino y el nombre de la hoja única que quieres crear.",
            style="Hint.TLabel",
            wraplength=580,
            justify="left",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 18))

        self._build_file_field(
            parent=form_card,
            row=2,
            label="Archivo Excel",
            variable=self.input_file_var,
            summary_var=self.input_summary_var,
            button_text="Examinar",
            button_command=self.select_input_file,
        )
        self._build_file_field(
            parent=form_card,
            row=4,
            label="Carpeta de salida",
            variable=self.output_dir_var,
            summary_var=self.output_summary_var,
            button_text="Elegir",
            button_command=self.select_output_dir,
        )
        self._build_sheet_field(form_card, row=6)

        log_card = ttk.Frame(content, style="Card.TFrame", padding=20)
        log_card.grid(row=0, column=1, sticky="nsew")
        log_card.columnconfigure(0, weight=1)
        log_card.rowconfigure(1, weight=1)

        ttk.Label(log_card, text="Seguimiento en tiempo real", style="FieldLabel.TLabel").grid(
            row=0, column=0, sticky="w"
        )

        text_frame = ttk.Frame(log_card, style="Card.TFrame")
        text_frame.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            text_frame,
            wrap="word",
            height=18,
            state="disabled",
            bg="#0f172a",
            fg="#e2e8f0",
            insertbackground="#e2e8f0",
            relief="flat",
            padx=12,
            pady=12,
            font=("Consolas", 10),
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        tips = ttk.Frame(log_card, style="Card.TFrame", padding=(0, 14, 0, 0))
        tips.grid(row=2, column=0, sticky="ew")
        tips.columnconfigure(0, weight=1)
        ttk.Label(
            tips,
            text=(
                "Tip: si no indicas carpeta de salida, al elegir el archivo se sugerirá una carpeta llamada 'salida'."
            ),
            style="Hint.TLabel",
            wraplength=360,
            justify="left",
        ).grid(row=0, column=0, sticky="w")

    def _build_file_field(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        summary_var: tk.StringVar,
        button_text: str,
        button_command,
    ) -> None:
        ttk.Label(parent, text=label, style="FieldLabel.TLabel").grid(row=row, column=0, sticky="w", pady=(0, 4))
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=10)
        ttk.Button(parent, text=button_text, style="Secondary.TButton", command=button_command).grid(
            row=row, column=2, sticky="ew"
        )
        ttk.Label(parent, textvariable=summary_var, style="Hint.TLabel", wraplength=580, justify="left").grid(
            row=row + 1, column=0, columnspan=3, sticky="w", pady=(4, 12)
        )

    def _build_sheet_field(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="Nombre de la hoja final", style="FieldLabel.TLabel").grid(
            row=row + 1, column=0, sticky="w", pady=(0, 4)
        )
        ttk.Entry(parent, textvariable=self.sheet_name_var).grid(row=row + 1, column=1, sticky="ew", padx=10)
        ttk.Button(
            parent,
            text="Restaurar AGRUPADOR",
            style="Secondary.TButton",
            command=self.restore_default_sheet_name,
        ).grid(row=row + 1, column=2, sticky="ew")
        ttk.Label(parent, textvariable=self.sheet_summary_var, style="Hint.TLabel", wraplength=580, justify="left").grid(
            row=row + 2, column=0, columnspan=3, sticky="w", pady=(4, 0)
        )

    def _build_footer(self) -> None:
        footer = ttk.Frame(self, style="App.TFrame", padding=(22, 0, 22, 22))
        footer.grid(row=3, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)

        ttk.Label(footer, textvariable=self.status_var, style="Status.TLabel").grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(footer, mode="indeterminate", style="Modern.Horizontal.TProgressbar", length=180)
        self.progress.grid(row=0, column=1, padx=12)
        self.process_button = ttk.Button(
            footer,
            text="Procesar archivo",
            style="Primary.TButton",
            command=self.process_file,
        )
        self.process_button.grid(row=0, column=2, sticky="e")

    def restore_default_sheet_name(self) -> None:
        self.sheet_name_var.set(DEFAULT_SHEET_NAME)
        self._refresh_summary()

    def select_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xlsm")],
        )
        if file_path:
            self.input_file_var.set(file_path)
            if not self.output_dir_var.get().strip():
                self.output_dir_var.set(str(Path(file_path).parent / "salida"))
            self._refresh_summary()

    def select_output_dir(self) -> None:
        folder_path = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if folder_path:
            self.output_dir_var.set(folder_path)
            self._refresh_summary()

    def append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.master.update_idletasks()

    def set_processing_state(self, processing: bool) -> None:
        state = "disabled" if processing else "normal"
        self.process_button.configure(state=state)
        if processing:
            self.progress.start(10)
        else:
            self.progress.stop()

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
        self.status_var.set("Procesando archivo, validando CUPS y preparando la salida...")
        self.append_log("=" * 70)
        self.append_log("Inicio del proceso")
        self.append_log(f"Hoja final solicitada: {self.sheet_name_var.get().strip()}")

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

            self.master.after(
                0,
                self._on_success,
                result.output_file,
                result.total_rows,
                result.processed_sheets,
                result.error_rows,
                result.total_input_rows,
            )
        except AgrupadorError as exc:
            self.master.after(0, self._on_error, str(exc))
        except Exception as exc:  # noqa: BLE001
            self.master.after(0, self._on_error, f"Error inesperado: {exc}")

    def _thread_safe_log(self, message: str) -> None:
        self.master.after(0, self.append_log, message)

    def _refresh_summary(self) -> None:
        input_path = self.input_file_var.get().strip()
        output_path = self.output_dir_var.get().strip()
        sheet_name = self.sheet_name_var.get().strip() or DEFAULT_SHEET_NAME

        if input_path:
            input_file = Path(input_path)
            self.input_summary_var.set(
                f"Archivo seleccionado: {input_file.name} • Ubicación: {input_file.parent}"
            )
        else:
            self.input_summary_var.set("Aún no has seleccionado un archivo.")

        if output_path:
            self.output_summary_var.set(f"El archivo procesado se guardará en: {output_path}")
        else:
            self.output_summary_var.set("Se creará una carpeta de salida cuando elijas un archivo.")

        self.sheet_summary_var.set(
            f"El Excel generado contendrá únicamente la hoja '{sheet_name}' con toda la información agrupada."
        )

    def _on_success(
        self,
        output_file: Path,
        total_rows: int,
        processed_sheets: int,
        error_rows: int,
        total_input_rows: int,
    ) -> None:
        self.append_log(f"Hojas procesadas: {processed_sheets}")
        self.append_log(f"Total filas leídas: {total_input_rows:,}")
        self.append_log(f"Filas válidas en {self.sheet_name_var.get().strip() or DEFAULT_SHEET_NAME}: {total_rows:,}")
        self.append_log(f"Filas enviadas a ERRORES: {error_rows:,}")
        self.append_log(f"Archivo de salida: {output_file}")
        self.append_log("Proceso finalizado correctamente.")
        self.status_var.set("Proceso completado correctamente.")
        self.set_processing_state(False)
        messagebox.showinfo(
            PROJECT_NAME,
            "Proceso completado correctamente.\n\n"
            f"Filas válidas: {total_rows:,}\n"
            f"Filas con error: {error_rows:,}\n\n"
            f"Archivo generado:\n{output_file}",
        )

    def _on_error(self, message: str) -> None:
        self.append_log(message)
        self.status_var.set("Se produjo un error durante el proceso.")
        self.set_processing_state(False)
        messagebox.showerror(PROJECT_NAME, message)
