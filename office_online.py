from __future__ import annotations

import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path


WINDOWS_APPS_WINGET = Path.home() / "AppData" / "Local" / "Microsoft" / "WindowsApps" / "winget.exe"


@dataclass(frozen=True)
class OfficeOnlinePackage:
    action_id: str
    label: str
    winget_id: str


@dataclass
class OfficeOnlineStatus:
    available: bool
    message: str


OFFICE_ONLINE_PACKAGES: dict[str, OfficeOnlinePackage] = {
    "online_office_2024_proplus": OfficeOnlinePackage("online_office_2024_proplus", "Office Professional Plus 2024", "Microsoft.Office.ProfessionalPlus.2024"),
    "online_office_2024_home_business": OfficeOnlinePackage("online_office_2024_home_business", "Office Home & Business 2024", "Microsoft.Office.HomeBusiness.2024"),
    "online_office_2021_proplus": OfficeOnlinePackage("online_office_2021_proplus", "Office Professional Plus 2021", "Microsoft.Office.ProfessionalPlus.2021"),
    "online_office_2021_home_student": OfficeOnlinePackage("online_office_2021_home_student", "Office Home & Student 2021", "Microsoft.Office.HomeStudent.2021"),
    "online_microsoft_365": OfficeOnlinePackage("online_microsoft_365", "Microsoft 365", "Microsoft.Office"),
    "online_office_2019_proplus": OfficeOnlinePackage("online_office_2019_proplus", "Office Professional Plus 2019", "Microsoft.Office.ProfessionalPlus.2019"),
    "online_office_2016_proplus": OfficeOnlinePackage("online_office_2016_proplus", "Office Professional Plus 2016", "Microsoft.Office.ProfessionalPlus.2016"),
    "online_office_2013_proplus": OfficeOnlinePackage("online_office_2013_proplus", "Office Professional Plus 2013", "Microsoft.Office.ProfessionalPlus.2013"),
    "online_visio_2024_pro": OfficeOnlinePackage("online_visio_2024_pro", "Visio Professional 2024", "Microsoft.Visio.Professional.2024"),
    "online_project_2024_pro": OfficeOnlinePackage("online_project_2024_pro", "Project Professional 2024", "Microsoft.Project.Professional.2024"),
    "online_visio_2021_pro": OfficeOnlinePackage("online_visio_2021_pro", "Visio Professional 2021", "Microsoft.Visio.Professional.2021"),
    "online_project_2021_pro": OfficeOnlinePackage("online_project_2021_pro", "Project Professional 2021", "Microsoft.Project.Professional.2021"),
}


def get_online_package(action_id: str) -> OfficeOnlinePackage:
    return OFFICE_ONLINE_PACKAGES[action_id]


def find_winget_executable() -> str | None:
    path_candidate = shutil.which("winget")
    if path_candidate:
        return path_candidate
    try:
        if WINDOWS_APPS_WINGET.exists():
            return str(WINDOWS_APPS_WINGET)
    except OSError:
        return None
    return None


def check_online_package(action_id: str) -> OfficeOnlineStatus:
    package = get_online_package(action_id)
    winget_exe = find_winget_executable()
    if not winget_exe:
        return OfficeOnlineStatus(
            available=False,
            message="Winget не е открит на тази система. Online инсталацията не е налична.",
        )

    result = subprocess.run(
        [winget_exe, "show", "--id", package.winget_id, "--source", "winget"],
        capture_output=True,
        text=True,
        check=False,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if result.returncode == 0:
        return OfficeOnlineStatus(
            available=True,
            message=f"Налично за online инсталация чрез winget: {package.winget_id}",
        )

    error_text = (result.stderr or result.stdout or "").strip()
    if error_text:
        return OfficeOnlineStatus(
            available=False,
            message=f"Не е налично за online инсталация: {error_text}",
        )
    return OfficeOnlineStatus(
        available=False,
        message="Не е налично за online инсталация. Провери winget и package ID.",
    )
