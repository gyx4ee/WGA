from __future__ import annotations

import sys
from pathlib import Path


def project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def runtime_file(relative_path: str) -> Path:
    root = project_root()
    candidate = root / relative_path
    if candidate.exists():
        return candidate
    if getattr(sys, "_MEIPASS", None):
        bundled = Path(sys._MEIPASS) / relative_path  # type: ignore[attr-defined]
        if bundled.exists():
            return bundled
    return candidate


def ui_assets_root() -> Path:
    root = project_root()
    candidates = (
        root / "assets" / "wga_ui_v3",
        root / "WGA_UI_Assets_v3_COMPLETE" / "WGA_UI_Assets_v3",
        root / "WGA_UI_Assets_v3",
    )
    for candidate in candidates:
        if (candidate / "MANIFEST_V3.txt").exists():
            return candidate
    return candidates[0]


def ui_asset(relative_path: str) -> Path:
    return ui_assets_root() / relative_path


def loading_asset(relative_path: str) -> Path:
    root = project_root()
    preferred = root / "assets" / "loading_screen" / relative_path
    if preferred.exists():
        return preferred
    return runtime_file(f"assets/loading/{relative_path}")


def start_menu_asset(relative_path: str) -> Path:
    return runtime_file(f"assets/start_menu/{relative_path}")


def qss_path() -> Path:
    return runtime_file("ui/styles.qss")
