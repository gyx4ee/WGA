from __future__ import annotations

import ctypes
import json
import os
import platform
import subprocess
import sys
import tempfile
import winreg
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from collections.abc import Callable


PROFILE_DIR = Path(os.environ.get("LOCALAPPDATA", tempfile.gettempdir())) / "WGA" / "optimization"
BACKUP_FILE = PROFILE_DIR / "windows10_low_ram_backup.json"
OPEN_SHELL_BACKUP = PROFILE_DIR / "openshell_before_low_ram.reg"


@dataclass(frozen=True)
class RegistryChange:
    root: int
    root_name: str
    path: str
    name: str
    value: int | str
    value_type: int
    description: str


CHANGES = (
    RegistryChange(winreg.HKEY_CURRENT_USER, "HKCU", r"Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects", "VisualFXSetting", 2, winreg.REG_DWORD, "Визуални ефекти: най-добра производителност"),
    RegistryChange(winreg.HKEY_CURRENT_USER, "HKCU", r"Control Panel\Desktop\WindowMetrics", "MinAnimate", "0", winreg.REG_SZ, "Изключване на анимацията при минимизиране"),
    RegistryChange(winreg.HKEY_CURRENT_USER, "HKCU", r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "EnableTransparency", 0, winreg.REG_DWORD, "Изключване на прозрачността"),
    RegistryChange(winreg.HKEY_CURRENT_USER, "HKCU", r"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "TaskbarAnimations", 0, winreg.REG_DWORD, "Изключване на анимациите в лентата със задачи"),
    RegistryChange(winreg.HKEY_CURRENT_USER, "HKCU", r"Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager", "SoftLandingEnabled", 0, winreg.REG_DWORD, "Изключване на съветите и предложенията"),
    RegistryChange(winreg.HKEY_CURRENT_USER, "HKCU", r"Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager", "SubscribedContent-338389Enabled", 0, winreg.REG_DWORD, "Изключване на Windows Spotlight предложенията"),
    RegistryChange(winreg.HKEY_CURRENT_USER, "HKCU", r"Software\Microsoft\Windows\CurrentVersion\GameDVR", "AppCaptureEnabled", 0, winreg.REG_DWORD, "Изключване на фоновия Game DVR запис"),
    RegistryChange(winreg.HKEY_CURRENT_USER, "HKCU", r"System\GameConfigStore", "GameDVR_Enabled", 0, winreg.REG_DWORD, "Изключване на Game DVR"),
)


def profile_descriptions() -> list[str]:
    return [
        "Open-Shell меню в класически XP стил (ако Open-Shell е инсталиран)",
        *(change.description for change in CHANGES),
        "Създаване на архив за връщане на всички промени",
        "Опит за създаване на Windows Restore Point",
    ]


def is_windows_10() -> bool:
    if platform.system() != "Windows":
        return False
    try:
        return int(platform.version().split(".")[0]) == 10 and int(platform.version().split(".")[2]) < 22000
    except (ValueError, IndexError):
        return False


def _run(command: list[str], timeout: int = 45) -> subprocess.CompletedProcess[str]:
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    return subprocess.run(command, capture_output=True, text=True, timeout=timeout, creationflags=flags)


def find_open_shell() -> Path | None:
    bundle_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    candidates = (
        bundle_root / "third_party" / "open-shell" / "PFiles" / "Open-Shell" / "StartMenu.exe",
        bundle_root / "third_party" / "open-shell" / "portable" / "PFiles" / "Open-Shell" / "StartMenu.exe",
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Open-Shell" / "StartMenu.exe",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Open-Shell" / "StartMenu.exe",
    )
    return next((path for path in candidates if path.is_file()), None)


def _read_existing(change: RegistryChange) -> dict[str, object]:
    try:
        with winreg.OpenKey(change.root, change.path) as key:
            value, value_type = winreg.QueryValueEx(key, change.name)
        return {"exists": True, "value": value, "type": value_type}
    except FileNotFoundError:
        return {"exists": False}


def _create_backup() -> dict[str, object]:
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    backup = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "registry": {
            f"{item.root_name}\\{item.path}|{item.name}": _read_existing(item)
            for item in CHANGES
        },
    }
    BACKUP_FILE.write_text(json.dumps(backup, ensure_ascii=False, indent=2), encoding="utf-8")
    if find_open_shell():
        _run(["reg.exe", "export", r"HKCU\Software\OpenShell\StartMenu", str(OPEN_SHELL_BACKUP), "/y"])
    return backup


def _create_restore_point() -> str:
    command = (
        "Checkpoint-Computer -Description 'WGA before Windows 10 low RAM profile' "
        "-RestorePointType MODIFY_SETTINGS -ErrorAction Stop"
    )
    result = _run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command], 90)
    if result.returncode == 0:
        return "Restore Point е създаден."
    return "Restore Point не бе създаден (Windows допуска най-много един на 24 часа или защитата е изключена)."


def _apply_open_shell(executable: Path) -> str:
    xml_path = PROFILE_DIR / "openshell_xp_low_ram.xml"
    xml_path.write_text(
        """<?xml version="1.0" encoding="utf-8"?>
<Settings component="StartMenu" version="4.4.190">
  <MenuStyle value="Classic2"/>
  <SkinC2 value="Windows XP Luna"/>
  <EnableGlass value="0"/>
  <MenuDelay value="100"/>
  <SplitMenuDelay value="100"/>
  <MainMenuAnimation value="None"/>
  <SubMenuAnimation value="None"/>
  <ShowUserName value="1"/>
  <ShowUserPicture value="1"/>
  <RecentPrograms value="Recent"/>
  <MaxRecentPrograms value="8"/>
  <SearchBox value="Normal"/>
</Settings>
""",
        encoding="utf-8",
    )
    imported = _run([str(executable), "-xml", str(xml_path)])
    if imported.returncode != 0:
        raise RuntimeError(imported.stderr.strip() or "Open-Shell не прие XML настройките.")
    _run([str(executable), "-reloadsettings"])
    return "Open-Shell е настроен с двуколонно класическо XP оформление."


ProgressCallback = Callable[[int, int, str], None]


def _report(callback: ProgressCallback | None, step: int, total: int, message: str) -> None:
    if callback:
        callback(step, total, message)


def apply_profile(progress_callback: ProgressCallback | None = None) -> list[str]:
    if not is_windows_10():
        raise RuntimeError("Този профил може да се приложи само на Windows 10.")
    if BACKUP_FILE.exists():
        raise RuntimeError("Профилът вече е прилаган. Първо използвайте „Върни настройките“.")

    total_steps = len(CHANGES) + 5
    _report(progress_callback, 1, total_steps, "Създаване на архив на текущите настройки...")
    _create_backup()
    _report(progress_callback, 2, total_steps, "Създаване на Windows Restore Point...")
    messages = [_create_restore_point()]
    try:
        for index, change in enumerate(CHANGES, start=3):
            _report(progress_callback, index, total_steps, change.description)
            with winreg.CreateKeyEx(change.root, change.path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, change.name, 0, change.value_type, change.value)
        _report(progress_callback, len(CHANGES) + 3, total_steps, "Настройване на Open-Shell XP менюто...")
        open_shell = find_open_shell()
        if open_shell:
            messages.append(_apply_open_shell(open_shell))
        else:
            messages.append("Open-Shell не е инсталиран. Оптимизациите са приложени, но XP менюто е пропуснато.")
        _report(progress_callback, len(CHANGES) + 4, total_steps, "Прилагане на промените в Windows Explorer...")
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x001A, 0, "Environment", 0x0002, 5000, None)
        messages.append("Профилът за 2/4 GB RAM е приложен. Излезте и влезте отново в Windows за всички визуални промени.")
        _report(progress_callback, total_steps, total_steps, "Windows 10 оптимизацията приключи успешно.")
        return messages
    except Exception:
        restore_profile()
        raise


def restore_profile(progress_callback: ProgressCallback | None = None) -> list[str]:
    if not BACKUP_FILE.exists():
        raise RuntimeError("Няма намерен архив от приложен профил.")
    total_steps = len(CHANGES) + 3
    _report(progress_callback, 1, total_steps, "Зареждане на архива с предишните настройки...")
    backup = json.loads(BACKUP_FILE.read_text(encoding="utf-8"))
    registry_backup = backup.get("registry", {})
    for index, change in enumerate(CHANGES, start=2):
        _report(progress_callback, index, total_steps, f"Възстановяване: {change.description}")
        saved = registry_backup.get(f"{change.root_name}\\{change.path}|{change.name}", {"exists": False})
        with winreg.CreateKeyEx(change.root, change.path, 0, winreg.KEY_SET_VALUE) as key:
            if saved.get("exists"):
                winreg.SetValueEx(key, change.name, 0, int(saved["type"]), saved["value"])
            else:
                try:
                    winreg.DeleteValue(key, change.name)
                except FileNotFoundError:
                    pass
    _report(progress_callback, len(CHANGES) + 2, total_steps, "Възстановяване на Open-Shell и обновяване на Explorer...")
    if OPEN_SHELL_BACKUP.exists():
        _run(["reg.exe", "import", str(OPEN_SHELL_BACKUP)])
        executable = find_open_shell()
        if executable:
            _run([str(executable), "-reloadsettings"])
    BACKUP_FILE.unlink(missing_ok=True)
    ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x001A, 0, "Environment", 0x0002, 5000, None)
    _report(progress_callback, total_steps, total_steps, "Предишните Windows 10 настройки са възстановени.")
    return ["Предишните Windows и Open-Shell настройки са възстановени.", "Излезте и влезте отново в Windows, за да се обнови целият интерфейс."]
