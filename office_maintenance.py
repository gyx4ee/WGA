from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from office_online import find_winget_executable


@dataclass
class OfficeMaintenanceStatus:
    available: bool
    message: str


OSPP_SEARCH_ROOTS = (
    Path(r"C:\Program Files\Microsoft Office"),
    Path(r"C:\Program Files (x86)\Microsoft Office"),
)

CLICK_TO_RUN_CANDIDATES = (
    Path(r"C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe"),
    Path(r"C:\Program Files (x86)\Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe"),
)

OFFICE_FORCE_UNINSTALL_IDS = (
    "Microsoft.Office",
    "Microsoft.Office.ProfessionalPlus.2021",
    "Microsoft.Office.ProfessionalPlus.2019",
    "Microsoft.Office.ProfessionalPlus.2016",
    "Microsoft.Office.ProfessionalPlus.2013",
)


def find_ospp_vbs() -> Path | None:
    for search_root in OSPP_SEARCH_ROOTS:
        if not search_root.exists():
            continue
        for candidate in search_root.rglob("OSPP.VBS"):
            if candidate.is_file():
                return candidate
    return None


def find_click_to_run_executable() -> Path | None:
    for candidate in CLICK_TO_RUN_CANDIDATES:
        if candidate.exists():
            return candidate
    return None


def check_maintenance_action(action_id: str) -> OfficeMaintenanceStatus:
    if action_id == "office_check_activation_status":
        ospp_vbs = find_ospp_vbs()
        if ospp_vbs:
            return OfficeMaintenanceStatus(True, f"OSPP.VBS found: {ospp_vbs}")
        return OfficeMaintenanceStatus(False, "OSPP.VBS was not found. Install Office first.")

    if action_id == "office_quick_repair":
        click_to_run = find_click_to_run_executable()
        if click_to_run:
            return OfficeMaintenanceStatus(True, f"Repair tool found: {click_to_run}")
        return OfficeMaintenanceStatus(False, "OfficeClickToRun.exe was not found. Quick Repair is unavailable.")

    if action_id == "office_force_uninstall_all":
        winget_exe = find_winget_executable()
        if winget_exe:
            return OfficeMaintenanceStatus(True, f"Winget is available: {winget_exe}")
        return OfficeMaintenanceStatus(False, "Winget was not found. Force uninstall cannot start.")

    return OfficeMaintenanceStatus(False, "This maintenance action is not configured.")
