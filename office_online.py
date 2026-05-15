from __future__ import annotations

import shutil
from dataclasses import dataclass
from pathlib import Path


WINDOWS_APPS_WINGET = Path.home() / "AppData" / "Local" / "Microsoft" / "WindowsApps" / "winget.exe"
ODT_CONFIRMATION_URL = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"


@dataclass(frozen=True)
class OfficeOnlinePackage:
    # Opisva kak tochen Office paket se instalira online prez Microsoft ODT.
    action_id: str
    label: str
    product_id: str
    channel: str
    registry_patterns: tuple[str, ...]
    supported: bool = True
    support_note: str = ""

    @property
    def winget_id(self) -> str:
        # Zapazva starite mesta v UI, koito vse oshte pokazvat poleto winget_id.
        return self.product_id


@dataclass
class OfficeOnlineStatus:
    # Pazi rezultata ot proverkata dali online paketat moje da se startira.
    available: bool
    message: str


OFFICE_ONLINE_PACKAGES: dict[str, OfficeOnlinePackage] = {
    "online_office_2024_proplus": OfficeOnlinePackage(
        "online_office_2024_proplus",
        "Office Professional Plus 2024",
        "ProPlus2024Volume",
        "PerpetualVL2024",
        ("Office.*2024", "Professional Plus 2024"),
    ),
    "online_office_2024_home_business": OfficeOnlinePackage(
        "online_office_2024_home_business",
        "Office Home & Business 2024",
        "HomeBusiness2024Retail",
        "Current",
        ("Office.*2024", "Home.*Business 2024"),
    ),
    "online_office_2021_proplus": OfficeOnlinePackage(
        "online_office_2021_proplus",
        "Office Professional Plus 2021",
        "ProPlus2021Volume",
        "PerpetualVL2021",
        ("Office.*2021", "Professional Plus 2021"),
    ),
    "online_office_2021_home_student": OfficeOnlinePackage(
        "online_office_2021_home_student",
        "Office Home & Student 2021",
        "HomeStudent2021Retail",
        "Current",
        ("Office.*2021", "Home.*Student 2021"),
    ),
    "online_microsoft_365": OfficeOnlinePackage(
        "online_microsoft_365",
        "Microsoft 365",
        "O365ProPlusRetail",
        "MonthlyEnterprise",
        ("Microsoft 365", "Office 365"),
    ),
    "online_office_2019_proplus": OfficeOnlinePackage(
        "online_office_2019_proplus",
        "Office Professional Plus 2019",
        "ProPlus2019Volume",
        "PerpetualVL2019",
        ("Office.*2019", "Professional Plus 2019"),
    ),
    "online_office_2016_proplus": OfficeOnlinePackage(
        "online_office_2016_proplus",
        "Office Professional Plus 2016",
        "",
        "",
        ("Office.*2016", "Professional Plus 2016"),
        supported=False,
        support_note="Office Professional Plus 2016 veche ne se poddarzha za nova online instalaciya prez sashtia ODT potok.",
    ),
    "online_office_2013_proplus": OfficeOnlinePackage(
        "online_office_2013_proplus",
        "Office Professional Plus 2013",
        "",
        "",
        ("Office.*2013", "Professional Plus 2013"),
        supported=False,
        support_note="Office 2013 ne e poddarzhan za tozi online installer i trqbva da se polzva drug metod.",
    ),
    "online_visio_2024_pro": OfficeOnlinePackage(
        "online_visio_2024_pro",
        "Visio Professional 2024",
        "VisioPro2024Volume",
        "PerpetualVL2024",
        ("Visio.*2024", "Visio Professional 2024"),
    ),
    "online_project_2024_pro": OfficeOnlinePackage(
        "online_project_2024_pro",
        "Project Professional 2024",
        "ProjectPro2024Volume",
        "PerpetualVL2024",
        ("Project.*2024", "Project Professional 2024"),
    ),
    "online_visio_2021_pro": OfficeOnlinePackage(
        "online_visio_2021_pro",
        "Visio Professional 2021",
        "VisioPro2021Volume",
        "PerpetualVL2021",
        ("Visio.*2021", "Visio Professional 2021"),
    ),
    "online_project_2021_pro": OfficeOnlinePackage(
        "online_project_2021_pro",
        "Project Professional 2021",
        "ProjectPro2021Volume",
        "PerpetualVL2021",
        ("Project.*2021", "Project Professional 2021"),
    ),
}


def get_online_package(action_id: str) -> OfficeOnlinePackage:
    # Vrashta dannite za konkretniya online Office paket.
    return OFFICE_ONLINE_PACKAGES[action_id]


def find_winget_executable() -> str | None:
    # Tarsi winget kakto v PATH, taka i v WindowsApps.
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
    # Proverkata veche e za ODT poddrzhka, a ne za star nevaliden winget paket.
    package = get_online_package(action_id)
    if not package.supported or not package.product_id or not package.channel:
        return OfficeOnlineStatus(
            available=False,
            message=package.support_note or "Tozi online Office paket ne se poddarzha v teku6tata konfiguraciya.",
        )

    return OfficeOnlineStatus(
        available=True,
        message=(
            "Paketat e gotov za online instalaciya prez Microsoft Office Deployment Tool. "
            "Nujen e internet i administratorski prava."
        ),
    )
