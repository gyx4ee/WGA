from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from services.paths import project_root


LEGACY_MENU_MAP = {
    "windows": "activation",
    "windows10": "windows10_activation",
    "windows11": "windows11_activation",
    "office": "office_activation",
    "office_install": "office_install_center",
    "optimization": "main",
    "drivers": "driver_backup",
    "network": "main",
    "system": "main",
    "tools": "install_software",
    "personalization": "language",
    "privacy": "main",
    "advanced": "hidden_menu",
    "onedrive": "reset_onedrive",
    "auto_installer": "auto_installer",
}


def launch_legacy_menu(menu_key: str | None = None) -> subprocess.Popen:
    root = project_root()
    if getattr(sys, "frozen", False):
        command = [str(Path(sys.executable).resolve()), "--legacy"]
    else:
        command = [sys.executable, str(root / "legacy_app.py")]
    if menu_key:
        command.extend(["--menu", menu_key])
    return subprocess.Popen(command, cwd=str(root), close_fds=True)


def open_legacy_section(section_key: str) -> None:
    launch_legacy_menu(LEGACY_MENU_MAP.get(section_key, section_key))
