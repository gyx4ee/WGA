# Помощен модул за намиране на Office activation скрипта и подготовка на командите.
from __future__ import annotations

import os
from pathlib import Path


OFFICE_VERSION_LABELS = {
    "office_2016_activation": "Office 2016",
    "office_2019_activation": "Office 2019",
    "office_2021_activation": "Office 2021",
}


# Връща office version label в удобен за останалия код вид.
def get_office_version_label(action_id: str) -> str:
    # Свързва вътрешното action_id с човешкото име на Office версията.
    return OFFICE_VERSION_LABELS[action_id]


# Намира ospp script.
def locate_ospp_script(version_label: str) -> Path:
    # Търси OSPP.VBS, защото той прави Office активацията.
    office_folder = "Office16"
    candidate_roots = [
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Microsoft Office" / office_folder,
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Microsoft Office" / office_folder,
    ]

    for root in candidate_roots:
        candidate = root / "ospp.vbs"
        if candidate.exists():
            return candidate

    raise FileNotFoundError(
        f"{version_label} activation script was not found. Expected ospp.vbs under Microsoft Office\\{office_folder}."
    )


# Подготвя office activation commands според избраните настройки.
def build_office_activation_commands(version_label: str, product_key: str) -> list[tuple[int, str, list[str]]]:
    # Подготвя стъпките за въвеждане на ключ и после активация.
    ospp_script = locate_ospp_script(version_label)
    return [
        (
            45,
            f"Installing {version_label} product key...",
            ["cscript", "//nologo", str(ospp_script), f"/inpkey:{product_key}"],
        ),
        (
            90,
            f"Requesting {version_label} activation...",
            ["cscript", "//nologo", str(ospp_script), "/act"],
        ),
    ]
