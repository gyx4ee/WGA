from __future__ import annotations

import winreg
from dataclasses import dataclass


UNINSTALL_PATHS = [
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
    (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
]


@dataclass
class InstalledOfficeInfo:
    installed: bool
    display_name: str = ""
    uninstall_string: str = ""


OFFICE_MATCH_RULES = {
    "install_office_2016_offline": ["Office", "2016"],
    "install_office_2019_offline": ["Office", "2019"],
    "install_office_2021_offline": ["Office", "2021"],
    "install_office_2021_new_offline": ["Office", "2021"],
    "install_office_2024_prof_offline": ["Office", "2024"],
    "install_office_2024_standard_offline": ["Office", "2024"],
    "install_office_2021_standard_offline": ["Office", "2021"],
}


def _iter_uninstall_entries() -> list[dict[str, str]]:
    entries: list[dict[str, str]] = []
    for hive, path in UNINSTALL_PATHS:
        try:
            with winreg.OpenKey(hive, path) as root:
                count = winreg.QueryInfoKey(root)[0]
                for index in range(count):
                    try:
                        subkey_name = winreg.EnumKey(root, index)
                        with winreg.OpenKey(root, subkey_name) as subkey:
                            display_name = _query_value(subkey, "DisplayName")
                            uninstall_string = _query_value(subkey, "UninstallString")
                            if display_name:
                                entries.append(
                                    {
                                        "display_name": display_name,
                                        "uninstall_string": uninstall_string,
                                    }
                                )
                    except OSError:
                        continue
        except OSError:
            continue
    return entries


def _query_value(key: winreg.HKEYType, value_name: str) -> str:
    try:
        value, _ = winreg.QueryValueEx(key, value_name)
        return str(value).strip()
    except OSError:
        return ""


def detect_installed_office(action_id: str) -> InstalledOfficeInfo:
    match_parts = OFFICE_MATCH_RULES.get(action_id, [])
    for entry in _iter_uninstall_entries():
        display_name = entry["display_name"]
        lowered = display_name.lower()
        if all(part.lower() in lowered for part in match_parts):
            return InstalledOfficeInfo(
                installed=True,
                display_name=display_name,
                uninstall_string=entry["uninstall_string"],
            )
    return InstalledOfficeInfo(installed=False)
