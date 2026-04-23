from __future__ import annotations

import re
import subprocess
from dataclasses import dataclass
from pathlib import Path

from office_online import find_winget_executable
from path_utils import resolve_installers_root


ADOBE_READER_WINGET_ID = "Adobe.Acrobat.Reader.64-bit"
ADOBE_INSTALLER_PATTERNS = (
    "AcroRdr*.exe",
    "Reader*.exe",
    "AdobeReader*.exe",
    "AcrobatReader*.exe",
)


@dataclass(frozen=True)
class AdobeReaderStatus:
    installed_version: str
    latest_version: str
    winget_available: bool
    local_installer: Path | None
    local_installer_version: str
    message: str

    @property
    def has_local_installer(self) -> bool:
        return self.local_installer is not None

    @property
    def local_matches_latest(self) -> bool:
        return bool(
            self.local_installer_version
            and self.latest_version
            and self.local_installer_version == self.latest_version
        )


def _run_command(command: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        capture_output=True,
        text=True,
        check=False,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def _extract_winget_field(output: str, field: str) -> str:
    pattern = rf"^{re.escape(field)}\s*:\s*(.+)$"
    for line in output.splitlines():
        match = re.search(pattern, line.strip(), flags=re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return ""


def _latest_reader_version(winget_exe: str | None) -> tuple[str, str]:
    if not winget_exe:
        return "", "Winget не е открит. Онлайн проверката за актуална версия не е налична."
    result = _run_command([winget_exe, "show", "--id", ADOBE_READER_WINGET_ID, "--source", "winget"])
    output = f"{result.stdout}\n{result.stderr}".strip()
    if result.returncode != 0:
        return "", output or "Winget не успя да провери Adobe Reader."
    version = _extract_winget_field(output, "Version")
    return version, f"Актуална версия според winget: {version or 'неизвестна'}."


def _installed_reader_version(winget_exe: str | None) -> str:
    if not winget_exe:
        return ""
    result = _run_command([winget_exe, "list", "--id", ADOBE_READER_WINGET_ID, "--source", "winget"])
    output = f"{result.stdout}\n{result.stderr}".strip()
    if result.returncode != 0 or ADOBE_READER_WINGET_ID.lower() not in output.lower():
        return ""
    for line in output.splitlines():
        if ADOBE_READER_WINGET_ID.lower() in line.lower():
            parts = [part for part in re.split(r"\s{2,}", line.strip()) if part]
            if len(parts) >= 3:
                return parts[2]
    return "Installed"


def _file_version(path: Path) -> str:
    safe_path = str(path).replace("'", "''")
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        f"(Get-Item -LiteralPath '{safe_path}').VersionInfo.ProductVersion",
    ]
    result = _run_command(command)
    if result.returncode == 0:
        return result.stdout.strip()
    return ""


def find_local_adobe_installer(program_root: Path) -> Path | None:
    installers_root = resolve_installers_root(program_root)
    if not installers_root.exists():
        return None

    candidates: list[Path] = []
    preferred_dirs = [
        installers_root / "AdobeReader",
        installers_root / "Adobe Reader",
        installers_root / "Adobe",
        installers_root,
    ]
    for directory in preferred_dirs:
        if not directory.exists():
            continue
        for pattern in ADOBE_INSTALLER_PATTERNS:
            candidates.extend(directory.glob(pattern))

    if not candidates:
        for pattern in ADOBE_INSTALLER_PATTERNS:
            candidates.extend(installers_root.rglob(pattern))

    files = [candidate for candidate in candidates if candidate.is_file()]
    if not files:
        return None
    return max(files, key=lambda path: path.stat().st_mtime)


def check_adobe_reader_status(program_root: Path) -> AdobeReaderStatus:
    winget_exe = find_winget_executable()
    latest_version, latest_message = _latest_reader_version(winget_exe)
    installed_version = _installed_reader_version(winget_exe)
    local_installer = find_local_adobe_installer(program_root)
    local_version = _file_version(local_installer) if local_installer else ""

    if local_installer and latest_version and local_version and local_version != latest_version:
        message = (
            f"Локалният installer е версия {local_version}, а актуалната е {latest_version}. "
            "Препоръчително е да се използва онлайн инсталация или да се замени файлът."
        )
    elif local_installer:
        message = f"Локален installer: {local_installer}"
    else:
        message = "Локален Adobe Reader installer не е открит. Препоръчва се онлайн инсталация чрез winget."

    return AdobeReaderStatus(
        installed_version=installed_version,
        latest_version=latest_version,
        winget_available=winget_exe is not None,
        local_installer=local_installer,
        local_installer_version=local_version,
        message=f"{latest_message} {message}".strip(),
    )
