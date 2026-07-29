# Описва локалните Office installer пакети и техните configuration файлове.
from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path

from path_utils import resolve_installers_root


# Помощна функция за current project root.
def current_project_root() -> Path:
    # Връща папката на програмата, независимо дали е .py или .exe билд.
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


# Класът OfficeInstaller групира свързана логика и състояние.
@dataclass(frozen=True)
class OfficeInstaller:
    # Описва къде са setup и config файловете за дадена Office версия.
    action_id: str
    label: str
    folder: str
    config_name: str

    # Помощна функция за setup path.
    @property
    def setup_path(self) -> Path:
        # Това е setup.exe за избраната версия.
        return self.installers_root / self.folder / "setup.exe"

    # Помощна функция за config path.
    @property
    def config_path(self) -> Path:
        # Това е XML конфигурацията за silent/controlled install.
        return self.installers_root / self.folder / self.config_name

    # Помощна функция за installers root.
    @property
    def installers_root(self) -> Path:
        # Взимаме общата Installers папка за текущото място на програмата.
        return resolve_installers_root(current_project_root())


OFFICE_OFFLINE_INSTALLERS: dict[str, OfficeInstaller] = {
    "install_office_2016_offline": OfficeInstaller(
        action_id="install_office_2016_offline",
        label="Office 2016 Offline",
        folder="Office2016",
        config_name="Configuration.xml",
    ),
    "install_office_2019_offline": OfficeInstaller(
        action_id="install_office_2019_offline",
        label="Office 2019 Offline",
        folder="Office2019",
        config_name="Configuration.xml",
    ),
    "install_office_2021_offline": OfficeInstaller(
        action_id="install_office_2021_offline",
        label="Office 2021 Offline",
        folder="Office2021",
        config_name="Configuration.xml",
    ),
    "install_office_2021_new_offline": OfficeInstaller(
        action_id="install_office_2021_new_offline",
        label="Office Professional 2021 Offline",
        folder="Office prof 2021",
        config_name="Configuration.xml",
    ),
    "install_office_2024_prof_offline": OfficeInstaller(
        action_id="install_office_2024_prof_offline",
        label="Office Professional 2024 Offline",
        folder="Office 2024 Prof",
        config_name="ConfigurationProPlus2024EnBgx64.xml",
    ),
    "install_office_2024_standard_offline": OfficeInstaller(
        action_id="install_office_2024_standard_offline",
        label="Office Standard 2024 Offline",
        folder="Office 2024 Standart",
        config_name="Configuration.xml",
    ),
    "install_office_2021_standard_offline": OfficeInstaller(
        action_id="install_office_2021_standard_offline",
        label="Office Standard 2021 Offline",
        folder="Office 2021 Standart",
        config_name="Configuration.xml",
    ),
}


# Връща office offline installer в удобен за останалия код вид.
def get_office_offline_installer(action_id: str) -> OfficeInstaller:
    # Връща готова конфигурация за избрания Office installer.
    return OFFICE_OFFLINE_INSTALLERS[action_id]
