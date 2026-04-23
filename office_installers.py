from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path

from path_utils import resolve_installers_root


def current_project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


@dataclass(frozen=True)
class OfficeInstaller:
    action_id: str
    label: str
    folder: str
    config_name: str

    @property
    def setup_path(self) -> Path:
        return self.installers_root / self.folder / "setup.exe"

    @property
    def config_path(self) -> Path:
        return self.installers_root / self.folder / self.config_name

    @property
    def installers_root(self) -> Path:
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


def get_office_offline_installer(action_id: str) -> OfficeInstaller:
    return OFFICE_OFFLINE_INSTALLERS[action_id]
