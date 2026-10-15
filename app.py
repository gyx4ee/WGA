from __future__ import annotations

import os
import sys
from pathlib import Path


if getattr(sys, "frozen", False) and getattr(sys, "_MEIPASS", None):
    pyside_dir = Path(sys._MEIPASS) / "PySide6"  # type: ignore[attr-defined]
    if pyside_dir.exists():
        os.environ["PATH"] = f"{pyside_dir}{os.pathsep}{os.environ.get('PATH', '')}"
        if hasattr(os, "add_dll_directory"):
            os.add_dll_directory(str(pyside_dir))

from PySide6.QtWidgets import QApplication, QMessageBox

from services.admin import is_running_as_admin, relaunch_as_admin


def run_tk_module(module_id: str) -> None:
    import tkinter as tk

    if module_id == "network":
        from network_monitoring_ui import NetworkMonitoringUI

        root = tk.Tk()
        NetworkMonitoringUI(root, on_back=root.destroy)
    elif module_id == "optimization":
        from optimization_ui import OptimizationUI

        root = tk.Tk()
        OptimizationUI(root, on_back=root.destroy)
    else:
        raise ValueError(f"Unknown module: {module_id}")
    root.mainloop()


def main() -> None:
    if "--network" in sys.argv:
        run_tk_module("network")
        return
    if "--optimization" in sys.argv:
        run_tk_module("optimization")
        return
    if "--legacy" in sys.argv:
        import legacy_app

        sys.argv = [sys.argv[0], *[arg for arg in sys.argv[1:] if arg != "--legacy"]]
        legacy_app.main()
        return

    app = QApplication(sys.argv)
    app.setApplicationName("WinSys Guardian Advanced")
    if not is_running_as_admin():
        if relaunch_as_admin():
            return
        QMessageBox.critical(None, "Administrator Rights Required", "WinSys Guardian Advanced must run with administrator rights.")
        return

    from ui.main_window import MainWindow

    window = MainWindow()
    window.resize(1240, 760)
    screen = app.primaryScreen()
    if screen is not None:
        available = screen.availableGeometry()
        frame = window.frameGeometry()
        frame.moveCenter(available.center())
        window.move(frame.topLeft())
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
