from __future__ import annotations

import os
from pathlib import Path


OFFICE_VERSION_LABELS = {
    "office_2016_activation": "Office 2016",
    "office_2019_activation": "Office 2019",
    "office_2021_activation": "Office 2021",
}


def get_office_version_label(action_id: str) -> str:
    return OFFICE_VERSION_LABELS[action_id]


def locate_ospp_script(version_label: str) -> Path:
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


def build_office_activation_commands(version_label: str, product_key: str) -> list[tuple[int, str, list[str]]]:
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
