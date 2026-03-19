from __future__ import annotations

import tkinter as tk

from validador_tarifario.ui.main_window import MainWindow
from validador_tarifario.utils.paths import PROJECT_NAME


def main() -> None:
    root = tk.Tk()
    root.title(PROJECT_NAME)
    root.geometry("1180x760")
    root.minsize(980, 640)
    MainWindow(root)
    root.mainloop()
