from __future__ import annotations

import shutil
import subprocess
from dataclasses import dataclass


@dataclass
class NexusToolStatus:
    available: bool
    message: str


def _run(command: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        capture_output=True,
        text=True,
        check=False,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def check_nexus_admin_status() -> NexusToolStatus:
    net_exe = shutil.which("net")
    powershell_exe = shutil.which("powershell")
    if not net_exe:
        return NexusToolStatus(False, "net.exe was not found. Local account actions are unavailable.")
    if not powershell_exe:
        return NexusToolStatus(False, "PowerShell was not found. User inspection features are unavailable.")
    return NexusToolStatus(True, f"Admin tools ready: {net_exe} and {powershell_exe}")


def list_users() -> subprocess.CompletedProcess[str]:
    script = (
        "Get-LocalUser | "
        "Select-Object Name,Enabled,Description,LastLogon | "
        "Sort-Object Name | Format-Table -AutoSize"
    )
    return _run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script])


def user_details(username: str) -> subprocess.CompletedProcess[str]:
    return _run(["net", "user", username])


def create_user(username: str, password: str | None, make_admin: bool) -> list[subprocess.CompletedProcess[str]]:
    results: list[subprocess.CompletedProcess[str]] = []
    if password:
        results.append(_run(["net", "user", username, password, "/add"]))
    else:
        results.append(_run(["net", "user", username, "/add"]))
    if make_admin and results[-1].returncode == 0:
        results.append(_run(["net", "localgroup", "Administrators", username, "/add"]))
    return results


def change_password(username: str, password: str) -> subprocess.CompletedProcess[str]:
    return _run(["net", "user", username, password])


def delete_user(username: str) -> subprocess.CompletedProcess[str]:
    return _run(["net", "user", username, "/delete"])


def set_admin_rights(username: str, make_admin: bool) -> subprocess.CompletedProcess[str]:
    if make_admin:
        return _run(["net", "localgroup", "Administrators", username, "/add"])
    return _run(["net", "localgroup", "Administrators", username, "/delete"])
