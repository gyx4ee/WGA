from __future__ import annotations

import ctypes
import json
import os
import platform
import subprocess
import tempfile
import winreg
from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


PROFILE_DIR = Path(os.environ.get("LOCALAPPDATA", tempfile.gettempdir())) / "WGA" / "optimization"
BACKUP_FILE = PROFILE_DIR / "windows11_safe_backup.json"
ProgressCallback = Callable[[int, int, str], None]


@dataclass(frozen=True)
class RegistryChange:
    root: int
    root_name: str
    path: str
    name: str
    value: int
    description: str


CHANGES = (
    RegistryChange(winreg.HKEY_CURRENT_USER, "HKCU", r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "EnableTransparency", 0, "Изключване на прозрачността"),
    RegistryChange(winreg.HKEY_CURRENT_USER, "HKCU", r"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "TaskbarAnimations", 0, "Изключване на анимациите в лентата със задачи"),
    RegistryChange(winreg.HKEY_CURRENT_USER, "HKCU", r"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "TaskbarDa", 0, "Изключване на Widgets от лентата със задачи"),
    RegistryChange(winreg.HKEY_CURRENT_USER, "HKCU", r"Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo", "Enabled", 0, "Изключване на advertising ID"),
    RegistryChange(winreg.HKEY_CURRENT_USER, "HKCU", r"Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager", "SoftLandingEnabled", 0, "Изключване на съветите и предложенията"),
    RegistryChange(winreg.HKEY_CURRENT_USER, "HKCU", r"Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager", "SubscribedContent-338389Enabled", 0, "Изключване на предложеното съдържание"),
    RegistryChange(winreg.HKEY_CURRENT_USER, "HKCU", r"Software\Microsoft\Windows\CurrentVersion\GameDVR", "AppCaptureEnabled", 0, "Изключване на фоновия Game DVR запис"),
    RegistryChange(winreg.HKEY_CURRENT_USER, "HKCU", r"System\GameConfigStore", "GameDVR_Enabled", 0, "Изключване на Game DVR"),
)


def profile_descriptions() -> list[str]:
    return [
        *(change.description for change in CHANGES),
        "Архив на всички променяни настройки",
        "Опит за създаване на Windows Restore Point",
    ]


def is_windows_11() -> bool:
    if platform.system() != "Windows":
        return False
    try:
        return int(platform.version().split(".")[2]) >= 22000
    except (ValueError, IndexError):
        return False


def _report(callback: ProgressCallback | None, step: int, total: int, message: str) -> None:
    if callback:
        callback(step, total, message)


def _run(command: list[str], timeout: int = 90) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        capture_output=True,
        text=True,
        timeout=timeout,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def _read_existing(change: RegistryChange) -> dict[str, object]:
    try:
        with winreg.OpenKey(change.root, change.path) as key:
            value, value_type = winreg.QueryValueEx(key, change.name)
        return {"exists": True, "value": value, "type": value_type}
    except FileNotFoundError:
        return {"exists": False}


def _create_backup() -> None:
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    registry = {
        f"{change.root_name}\\{change.path}|{change.name}": _read_existing(change)
        for change in CHANGES
    }
    BACKUP_FILE.write_text(
        json.dumps({"created_at": datetime.now().isoformat(timespec="seconds"), "registry": registry}, indent=2),
        encoding="utf-8",
    )


def _create_restore_point() -> str:
    command = "Checkpoint-Computer -Description 'WGA before Windows 11 safe profile' -RestorePointType MODIFY_SETTINGS -ErrorAction Stop"
    result = _run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command])
    if result.returncode == 0:
        return "Restore Point е създаден."
    return "Restore Point не бе създаден (защитата може да е изключена или вече има скорошна точка)."


def apply_profile(progress_callback: ProgressCallback | None = None) -> list[str]:
    if not is_windows_11():
        raise RuntimeError("Този профил може да се приложи само на Windows 11.")
    if BACKUP_FILE.exists():
        raise RuntimeError("Профилът вече е приложен. Първо използвайте „Върни настройките“.")

    total = len(CHANGES) + 4
    _report(progress_callback, 1, total, "Архивиране на текущите Windows 11 настройки...")
    _create_backup()
    _report(progress_callback, 2, total, "Създаване на Windows Restore Point...")
    messages = [_create_restore_point()]
    try:
        for index, change in enumerate(CHANGES, start=3):
            _report(progress_callback, index, total, change.description)
            with winreg.CreateKeyEx(change.root, change.path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, change.name, 0, winreg.REG_DWORD, change.value)
        _report(progress_callback, len(CHANGES) + 3, total, "Прилагане на промените в Windows Explorer...")
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x001A, 0, "Environment", 0x0002, 5000, None)
        messages.append("Безопасният Windows 11 профил е приложен. Излезте и влезте отново за всички визуални промени.")
        _report(progress_callback, total, total, "Windows 11 оптимизацията приключи успешно.")
        return messages
    except Exception:
        restore_profile()
        raise


def restore_profile(progress_callback: ProgressCallback | None = None) -> list[str]:
    if not BACKUP_FILE.exists():
        raise RuntimeError("Няма намерен архив от Windows 11 оптимизация.")
    total = len(CHANGES) + 3
    _report(progress_callback, 1, total, "Зареждане на архива с предишните настройки...")
    backup = json.loads(BACKUP_FILE.read_text(encoding="utf-8"))
    saved_registry = backup.get("registry", {})
    for index, change in enumerate(CHANGES, start=2):
        _report(progress_callback, index, total, f"Възстановяване: {change.description}")
        saved = saved_registry.get(f"{change.root_name}\\{change.path}|{change.name}", {"exists": False})
        with winreg.CreateKeyEx(change.root, change.path, 0, winreg.KEY_SET_VALUE) as key:
            if saved.get("exists"):
                winreg.SetValueEx(key, change.name, 0, int(saved["type"]), saved["value"])
            else:
                try:
                    winreg.DeleteValue(key, change.name)
                except FileNotFoundError:
                    pass
    _report(progress_callback, len(CHANGES) + 2, total, "Обновяване на Windows Explorer...")
    BACKUP_FILE.unlink(missing_ok=True)
    ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x001A, 0, "Environment", 0x0002, 5000, None)
    _report(progress_callback, total, total, "Предишните Windows 11 настройки са възстановени.")
    return ["Предишните Windows 11 настройки са възстановени.", "Излезте и влезте отново в Windows за пълно обновяване."]
