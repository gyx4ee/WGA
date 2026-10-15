from __future__ import annotations

import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.request
from queue import Queue
from threading import Thread
from contextvars import ContextVar
from collections.abc import Iterator
from datetime import datetime
from pathlib import Path

from adobe_reader import ADOBE_READER_WINGET_ID, check_adobe_reader_status
from driver_backup import compress_backup, create_backup_folder, create_driver_list, create_restore_script, desktop_path, export_drivers, generate_pc_report
from driver_backup import create_recovery_usb, detect_removable_drives, restore_drivers_from_backup
from language_manager import build_language_action, get_language_status
from legacy_app import enable_windows_desktop_icons
from nexus_admin import check_nexus_admin_status, list_users
from office_activation import build_office_activation_commands, get_office_version_label, locate_cscript_executable
from office_installers import OFFICE_OFFLINE_INSTALLERS, get_office_offline_installer
from office_maintenance import OFFICE_FORCE_UNINSTALL_IDS, check_maintenance_action, find_click_to_run_executable, find_ospp_vbs
from office_online import OFFICE_ONLINE_PACKAGES, ODT_CONFIRMATION_URL, check_online_package, get_online_package
from path_utils import resolve_installers_root
from resource_manager import check_resource_status, download_resource, missing_resource_report
from services.key_store import (
    WINDOWS_GENERIC_KEY_LABEL,
    normalize_windows_edition,
    office_legacy_store_key,
    office_key_label,
    office_store_key,
    windows_key_label,
    windows_store_key,
)

SECRET_INSTALL_PRESETS: dict[str, list[str]] = {
    "system_runtimes": ["Microsoft.DotNet.DesktopRuntime.8", "Microsoft.VCRedist.2015+.x64", "Microsoft.DirectX"],
    "browsers_comms": ["Google.Chrome", "Mozilla.Firefox", "Discord.Discord"],
    "development": ["Microsoft.VisualStudioCode", "Git.Git", "Docker.DockerDesktop"],
    "languages_db": ["Python.Python.3.13", "OpenJS.NodeJS.LTS", "Microsoft.SQLServerManagementStudio"],
    "cybersecurity_net": ["WiresharkFoundation.Wireshark", "Insecure.Nmap", "PuTTY.PuTTY"],
    "virtualization": ["Oracle.VirtualBox"],
    "multimedia_design": ["VideoLAN.VLC", "OBSProject.OBSStudio", "GIMP.GIMP"],
    "gaming_tools": ["Valve.Steam", "EpicGames.EpicGamesLauncher"],
    "utilities_office": ["7zip.7zip", "TheDocumentFoundation.LibreOffice", "AnyDeskSoftwareGmbH.AnyDesk"],
    "advanced_tools": ["Microsoft.Sysinternals", "voidtools.Everything"],
}

AUTO_INSTALLER_PACKAGES = [
    (".NET Runtime", "Microsoft.DotNet.Runtime.8"),
    (".NET Runtime 9", "Microsoft.DotNet.Runtime.9"),
    (".NET Windows Desktop Runtime", "Microsoft.DotNet.DesktopRuntime.8"),
    (".NET Windows Desktop Runtime 9", "Microsoft.DotNet.DesktopRuntime.9"),
    ("Google Chrome", "Google.Chrome"),
    ("Mozilla Firefox", "Mozilla.Firefox"),
    ("Microsoft Edge", "Microsoft.Edge"),
    ("Brave Browser", "Brave.Brave"),
    ("Opera Browser", "Opera.Opera"),
    ("Opera GX", "Opera.OperaGX"),
    ("Opera Air", "Opera.OperaAir"),
    ("Vivaldi Browser", "Vivaldi.Vivaldi"),
    ("LibreWolf Browser", "LibreWolf.LibreWolf"),
    ("Tor Browser", "TorProject.TorBrowser"),
    ("7-Zip", "7zip.7zip"),
    ("AnyDesk (платена)", "AnyDesk.AnyDesk"),
    ("ASP.NET Core Runtime 8", "Microsoft.DotNet.AspNetCore.8"),
    ("Audacity", "Audacity.Audacity"),
    ("DirectX Runtime", "Microsoft.DirectX"),
    ("Discord", "Discord.Discord"),
    ("Docker Desktop (платена)", "Docker.DockerDesktop"),
    ("Eclipse IDE for C/C++", "EclipseFoundation.Eclipse.CPP"),
    ("Eclipse IDE for Java", "EclipseFoundation.Eclipse.Java"),
    ("Everything Search", "voidtools.Everything"),
    ("Foxit PDF Reader", "Foxit.FoxitReader"),
    ("Git", "Git.Git"),
    ("GIMP", "GIMP.GIMP"),
    ("Google Chrome Remote Desktop", "Google.ChromeRemoteDesktopHost"),
    ("HopToDesk", "HopToDesk.HopToDesk"),
    ("Java Runtime", "EclipseAdoptium.Temurin.21.JRE"),
    ("JetBrains Toolbox (платена)", "JetBrains.Toolbox"),
    ("Microsoft Edge WebView2 Runtime", "Microsoft.EdgeWebView2Runtime"),
    ("Microsoft PowerToys", "Microsoft.PowerToys"),
    ("Microsoft Remote Desktop", "Microsoft.RemoteDesktopClient"),
    ("Microsoft SQL Server 2022 Developer", "Microsoft.SQLServer.2022.Developer"),
    ("Microsoft SQL Server Management Studio", "Microsoft.SQLServerManagementStudio"),
    ("Microsoft Teams", "Microsoft.Teams"),
    ("Microsoft Visual C++ Redistributable x64", "Microsoft.VCRedist.2015+.x64"),
    ("Microsoft Visual C++ Redistributable x86", "Microsoft.VCRedist.2015+.x86"),
    ("Microsoft Visual Studio 2022 Community (лиценз)", "Microsoft.VisualStudio.2022.Community"),
    ("NAPS2", "Cyanfish.NAPS2"),
    ("Nmap", "Insecure.Nmap"),
    ("Node.js LTS", "OpenJS.NodeJS.LTS"),
    ("Notepad++", "Notepad++.Notepad++"),
    ("OBS Studio", "OBSProject.OBSStudio"),
    ("Oracle Java Runtime (лиценз)", "Oracle.JavaRuntimeEnvironment"),
    ("Oracle SQL Developer", "Oracle.SQLDeveloper"),
    ("PDF24 Creator", "geeksoftwareGmbH.PDF24Creator"),
    ("PDFgear", "PDFgear.PDFgear"),
    ("PDFsam Basic", "PDFsam.PDFsam"),
    ("PDFtk Free", "PDFLabs.PDFtk.Free"),
    ("PDF-XChange Editor (платена)", "TrackerSoftware.PDF-XChangeEditor"),
    ("Postman", "Postman.Postman"),
    ("PuTTY", "PuTTY.PuTTY"),
    ("Python 3.13", "Python.Python.3.13"),
    ("Remote Desktop Connection Manager", "Microsoft.Sysinternals.RDCMan"),
    ("Slack", "SlackTechnologies.Slack"),
    ("SumatraPDF", "SumatraPDF.SumatraPDF"),
    ("TeamViewer (платена)", "TeamViewer.TeamViewer"),
    ("TeamViewer Host (платена)", "TeamViewer.TeamViewer.Host"),
    ("TeamViewer QuickSupport (платена)", "TeamViewer.TeamViewer.QuickSupport"),
    ("Telegram Desktop", "Telegram.TelegramDesktop"),
    ("UltraViewer", "DucFabulous.UltraViewer"),
    ("Visual Studio Code", "Microsoft.VisualStudioCode"),
    ("VLC Media Player", "VideoLAN.VLC"),
    ("Vulkan Runtime", "KhronosGroup.VulkanRT"),
    ("Windows App Runtime 1.8", "Microsoft.WindowsAppRuntime.1.8"),
    ("WinDirStat", "WinDirStat.WinDirStat"),
    ("WinRAR (платена)", "RARLab.WinRAR"),
    ("Wireshark", "WiresharkFoundation.Wireshark"),
    ("Zed Editor", "ZedIndustries.Zed"),
    ("Zoom", "Zoom.Zoom"),
    ("Steam", "Valve.Steam"),
]

LOCAL_INSTALLER_TASKS: dict[str, dict[str, str]] = {
    "install_ninite": {
        "label": "Ninite",
        "resource_id": "ninite_installer",
        "detect_mode": "none",
        "detect_value": "",
        "silent_args": "",
    },
    "install_visual_studio_setup": {
        "label": "Visual Studio Setup",
        "resource_id": "visual_studio_setup",
        "detect_mode": "winget",
        "detect_value": "Microsoft.VisualStudio.2022.Community",
        "silent_args": "",
    },
    "install_vscode_arm64": {
        "label": "VS Code ARM64",
        "resource_id": "vscode_user_setup_arm64",
        "detect_mode": "winget",
        "detect_value": "Microsoft.VisualStudioCode",
        "silent_args": "/VERYSILENT /NORESTART",
    },
}

AUTO_INSTALLER_LANGUAGE_PACKAGES = [
    ("Български езиков пакет", "toggle_bulgarian_language_pack"),
    ("Българска БДС клавиатура", "toggle_bulgarian_bds"),
    ("Българска фонетична клавиатура", "toggle_bulgarian_phonetic"),
    ("Традиционна фонетична клавиатура", "toggle_bulgarian_traditional"),
]

AUTO_INSTALLER_POST_INSTALL_ACTIONS = [
    ("Автоматична активация на Windows", "auto_windows_activation"),
    ("Автоматична активация на Office", "auto_office_activation"),
]

PREFERRED_FULL_PACKAGE_OFFICE_IDS = (
    "install_office_2024_prof_offline",
    "install_office_2024_standard_offline",
    "online_office_2024_proplus",
    "online_office_2024_home_business",
    "install_office_2021_new_offline",
    "install_office_2021_offline",
    "install_office_2021_standard_offline",
    "online_office_2021_proplus",
    "online_office_2021_home_student",
    "online_microsoft_365",
)

AUTO_INSTALLER_PRESETS: dict[str, dict[str, object]] = {
    "classroom": {
        "label": "Класна стая",
        "description": "Базови програми за учебна машина плюс Windows/Office активация.",
        "packages": {
            "Google.Chrome",
            "7zip.7zip",
            "RARLab.WinRAR",
            "VideoLAN.VLC",
            "Notepad++.Notepad++",
            "Microsoft.VCRedist.2015+.x64",
            "Microsoft.DirectX",
            "Microsoft.DotNet.DesktopRuntime.8",
            "EclipseAdoptium.Temurin.21.JRE",
            "voidtools.Everything",
            "resource_windirstat_tools",
            "install_adobe_reader",
            "install_office_2021_offline",
            "toggle_bulgarian_language_pack",
            "toggle_bulgarian_bds",
            "toggle_bulgarian_phonetic",
            "toggle_bulgarian_traditional",
            "auto_windows_activation",
            "auto_office_activation",
        },
    },
    "computer_lab": {
        "label": "Компютърен кабинет",
        "description": "Класна стая плюс инструменти за програмиране, разархивиране и подготовка на кабинет.",
        "packages": {
            "Google.Chrome",
            "7zip.7zip",
            "RARLab.WinRAR",
            "VideoLAN.VLC",
            "Notepad++.Notepad++",
            "Microsoft.VCRedist.2015+.x64",
            "Microsoft.DirectX",
            "Microsoft.DotNet.DesktopRuntime.8",
            "EclipseAdoptium.Temurin.21.JRE",
            "Git.Git",
            "Python.Python.3.13",
            "OpenJS.NodeJS.LTS",
            "Microsoft.VisualStudioCode",
            "install_visual_studio_setup",
            "install_vscode_arm64",
            "voidtools.Everything",
            "install_adobe_reader",
            "install_office_2021_offline",
            "toggle_bulgarian_language_pack",
            "toggle_bulgarian_bds",
            "toggle_bulgarian_phonetic",
            "toggle_bulgarian_traditional",
            "auto_windows_activation",
            "auto_office_activation",
        },
    },
    "full": {
        "label": "Пълен пакет",
        "description": "Маркира всички налични приложения, ресурси, езикови настройки и активации.",
        "packages": "all",
    },
}

_ACTION_CONTEXT: ContextVar[dict[str, str]] = ContextVar("wga_action_context", default={})


def run_action(action_id: str, title: str, mode: str = "action", selected_packages: list[str] | None = None) -> Iterator[dict[str, object]]:
    context_token = _ACTION_CONTEXT.set({"action_id": action_id, "title": title, "mode": mode})
    try:
        yield _event(3, "Подготовка", f"Стартира се: {title}")
        try:
            if action_id == "add_desktop_icons":
                yield from _run_add_desktop_icons()
            elif action_id == "reset_console":
                yield _event(100, "Готово", "Интерфейсът е опреснен. Използвай менюто отляво за навигация.", True)
            elif action_id in SECRET_INSTALL_PRESETS or action_id == "update_all_apps":
                yield from _run_secret_install_action(action_id)
            elif action_id in {"reset_onedrive_1", "reset_onedrive_2", "reset_onedrive_3"}:
                yield from _run_onedrive_reset(action_id)
            elif action_id == "install_adobe_reader":
                yield from _run_adobe_reader_install()
            elif action_id == "install_ninite":
                yield from _run_local_installer_task(action_id)
            elif action_id in LOCAL_INSTALLER_TASKS:
                yield from _run_local_installer_task(action_id)
            elif action_id.startswith("resource_"):
                yield from _run_generic_resource_action(action_id)
            elif action_id == "resources_status":
                yield from _run_resources_status()
            elif action_id == "resources_download_missing":
                yield from _run_resources_download_missing()
            elif action_id in OFFICE_OFFLINE_INSTALLERS:
                yield from _run_office_offline_install(action_id)
            elif action_id.startswith("online_"):
                yield from _run_office_online_install(action_id)
            elif action_id.startswith("nexus_"):
                yield from _run_nexus_action(action_id)
            elif mode == "installer" or action_id == "auto_installer":
                yield from _run_auto_installer(selected_packages)
            elif action_id == "auto_windows_activation":
                yield from _run_auto_windows_activation()
            elif action_id in {"activate_windows10", "activate_windows11"}:
                yield from _run_windows_activation(action_id)
            elif action_id == "auto_office_activation":
                yield from _run_auto_office_activation()
            elif action_id in {"office_2016_activation", "office_2019_activation", "office_2021_activation"}:
                yield from _run_office_activation(action_id)
            elif action_id == "office_check_activation_status":
                yield from _run_office_status()
            elif action_id == "office_quick_repair":
                yield from _run_office_repair_check()
            elif action_id == "office_force_uninstall_all":
                yield from _run_office_uninstall_check()
            elif action_id in {
                "language_refresh",
                "toggle_bulgarian_bds",
                "toggle_bulgarian_phonetic",
                "toggle_bulgarian_traditional",
                "toggle_bulgarian_language_pack",
                "remove_bulgarian_language",
            }:
                yield from _run_language_action(action_id)
            elif action_id in {"driver_backup_clean", "driver_backup_full", "driver_pc_report", "driver_recovery_usb", "driver_restore_last", "driver_backup_advanced"}:
                yield from _run_driver_action(action_id)
            elif action_id == "hidden_show_status":
                yield _event(100, "Готово", "Скритото меню работи и приложението е готово.", True)
            elif action_id == "hidden_load_agent_status":
                yield _event(100, "Готово", "Agent status модулът е наличен за свързване към новия интерфейс.", True)
            else:
                yield _event(
                    100,
                    "Непознато действие",
                    f"Действието '{action_id}' не съществува в регистъра на WGA backend действията. Операцията е спряна безопасно.",
                    False,
                )
        except Exception as exc:
            yield _event(100, "Грешка", str(exc), False)
    finally:
        _ACTION_CONTEXT.reset(context_token)


def is_registered_action(action_id: str, mode: str = "action") -> bool:
    direct_actions = {
        "add_desktop_icons",
        "reset_console",
        "install_adobe_reader",
        "install_ninite",
        "install_visual_studio_setup",
        "install_vscode_arm64",
        "auto_installer",
        "auto_windows_activation",
        "activate_windows10",
        "activate_windows11",
        "auto_office_activation",
        "office_2016_activation",
        "office_2019_activation",
        "office_2021_activation",
        "office_check_activation_status",
        "office_quick_repair",
        "office_force_uninstall_all",
        "language_refresh",
        "toggle_bulgarian_bds",
        "toggle_bulgarian_phonetic",
        "toggle_bulgarian_traditional",
        "toggle_bulgarian_language_pack",
        "remove_bulgarian_language",
        "driver_backup_clean",
        "driver_backup_full",
        "driver_pc_report",
        "driver_recovery_usb",
        "driver_restore_last",
        "driver_backup_advanced",
        "hidden_show_status",
        "hidden_load_agent_status",
        "update_all_apps",
        "resources_status",
        "resources_download_missing",
        "save_windows10_key",
        "show_windows10_key",
        "clear_windows10_key",
        "save_windows11_key",
        "show_windows11_key",
        "clear_windows11_key",
        "save_office_key",
        "show_office_key",
        "clear_office_key",
    }
    return (
        action_id in direct_actions
        or action_id in SECRET_INSTALL_PRESETS
        or action_id in OFFICE_OFFLINE_INSTALLERS
        or action_id.startswith("online_")
        or action_id.startswith("resource_")
        or action_id.startswith("nexus_")
        or action_id in {"reset_onedrive_1", "reset_onedrive_2", "reset_onedrive_3"}
        or mode == "installer"
    )


def _event(progress: int, status: str, detail: str, finished: bool = False) -> dict[str, object]:
    item = {"progress": progress, "status": status, "detail": detail, "finished": finished}
    _write_action_log(item)
    return item


def _write_action_log(item: dict[str, object]) -> None:
    try:
        import json

        log_dir = _project_root() / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        record = {"time": datetime.now().isoformat(timespec="seconds"), **_ACTION_CONTEXT.get(), **item}
        with (log_dir / "actions.jsonl").open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(record, ensure_ascii=False) + "\n")
    except OSError:
        return


def read_action_history(limit: int = 80) -> list[dict[str, object]]:
    import json

    log_path = _project_root() / "logs" / "actions.jsonl"
    if not log_path.exists():
        return []
    lines = log_path.read_text(encoding="utf-8", errors="replace").splitlines()[-limit:]
    history: list[dict[str, object]] = []
    for line in lines:
        try:
            record = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(record, dict):
            history.append(record)
    return history


def format_action_history(limit: int = 80) -> str:
    history = read_action_history(limit)
    if not history:
        return "Все още няма записани операции."
    blocks: list[str] = []
    for item in history:
        title = str(item.get("title") or item.get("action_id") or "WGA действие")
        status = str(item.get("status") or "")
        progress = str(item.get("progress") or "0")
        timestamp = str(item.get("time") or "")
        detail = str(item.get("detail") or "").strip()
        if len(detail) > 600:
            detail = detail[:600].rstrip() + "..."
        blocks.append(f"[{timestamp}] {title} ({progress}%) - {status}\n{detail}")
    return "\n\n".join(blocks)


def _run_command(command: list[str]) -> str:
    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        check=False,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    output = "\n".join(part.strip() for part in (result.stdout, result.stderr) if part and part.strip())
    if result.returncode != 0:
        raise RuntimeError(output or f"Command failed with code {result.returncode}.")
    return output or "Command completed."


def _project_root() -> Path:
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent.parent


def _run_command_result(command: list[str], cwd: Path | None = None, timeout: int | None = None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        capture_output=True,
        text=True,
        check=False,
        cwd=str(cwd) if cwd else None,
        timeout=timeout,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def _command_output(result: subprocess.CompletedProcess[str]) -> str:
    return "\n".join(part.strip() for part in (result.stdout, result.stderr) if part and part.strip())


def _secure_store() -> dict[str, str]:
    from legacy_app import load_secure_store

    return load_secure_store()


def _windows_script(script_name: str) -> Path:
    windir = Path(os.environ.get("WINDIR", r"C:\Windows"))
    candidates = [windir / "Sysnative" / script_name, windir / "System32" / script_name, windir / "SysWOW64" / script_name]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[1]


def _detect_windows_version() -> tuple[str, int, str, str]:
    if sys.platform != "win32":
        return "Windows", 0, WINDOWS_GENERIC_KEY_LABEL, platform.platform()
    version_info = sys.getwindowsversion()
    build = int(version_info.build)
    version_label = "Windows 11" if build >= 22000 else "Windows 10"
    edition_label = normalize_windows_edition(platform.win32_edition())
    release = platform.release()
    return version_label, build, edition_label, f"{release} {edition_label} build {build}"


def _saved_windows_key(version_label: str, edition_label: str) -> tuple[str, str, bool]:
    store = _secure_store()
    edition_store_key = windows_store_key(version_label, edition_label)
    product_key = store.get(edition_store_key, "").strip()
    if product_key:
        return product_key, edition_store_key, True
    generic_store_key = windows_store_key(version_label, WINDOWS_GENERIC_KEY_LABEL)
    return store.get(generic_store_key, "").strip(), generic_store_key, False


def _run_auto_windows_activation() -> Iterator[dict[str, object]]:
    version_label, build, edition_label, detected = _detect_windows_version()
    yield _event(18, "Проверка на Windows", f"Засечена система: {detected}.")
    if version_label not in {"Windows 10", "Windows 11"}:
        yield _event(100, "Неподдържана версия", f"Не може автоматично да се избере ключ за: {detected}.", False)
        return
    action_id = "activate_windows11" if version_label == "Windows 11" else "activate_windows10"
    yield _event(35, "Избор на лиценз", f"{version_label} {edition_label} build {build}: проверява се записан ключ за този edition.")
    product_key, store_key, exact_match = _saved_windows_key(version_label, edition_label)
    if not product_key:
        yield _event(100, "Липсва ключ", f"Няма записан ключ за {windows_key_label(version_label, edition_label)} или общ ключ за {version_label}.", False)
        return
    key_kind = windows_key_label(version_label, edition_label) if exact_match else windows_key_label(version_label, WINDOWS_GENERIC_KEY_LABEL)
    yield _event(48, "Ключът е намерен", f"Намерен е записан ключ: {key_kind}. Store: {store_key}. Ключът няма да се показва в логовете.")
    yield from _run_windows_activation(action_id, prechecked_key=product_key)


def _run_windows_activation(action_id: str, prechecked_key: str | None = None) -> Iterator[dict[str, object]]:
    version_label = "Windows 10" if action_id == "activate_windows10" else "Windows 11"
    _, _, edition_label, _ = _detect_windows_version()
    product_key = prechecked_key or _saved_windows_key(version_label, edition_label)[0]
    if not product_key:
        yield _event(100, "Липсва ключ", f"Няма записан продуктов ключ за {windows_key_label(version_label, edition_label)} или общ ключ за {version_label}.", False)
        return
    slmgr_path = _windows_script("slmgr.vbs")
    if not slmgr_path.exists():
        yield _event(100, "Липсва slmgr.vbs", f"Не е намерен Windows activation script: {slmgr_path}", False)
        return
    cscript = locate_cscript_executable()
    start_progress = 55 if prechecked_key else 20
    install_progress = 70 if prechecked_key else 45
    yield _event(start_progress, "Проверка", f"Намерен activation script: {slmgr_path}")
    yield _event(install_progress, "Инсталиране на ключ", _run_command([cscript, "//nologo", str(slmgr_path), "/ipk", product_key]))
    yield _event(90, "Активация", _run_command([cscript, "//nologo", str(slmgr_path), "/ato"]))
    yield _event(100, "Готово", f"{version_label} activation workflow завърши.", True)


def _run_office_activation(action_id: str) -> Iterator[dict[str, object]]:
    version_label = get_office_version_label(action_id)
    product_key, store_key = _saved_office_key(version_label)
    if not product_key:
        yield _event(100, "Липсва ключ", f"Няма записан Office ключ за {version_label}.", False)
        return
    yield _event(20, "Подготовка", f"Търси се OSPP.VBS за {version_label}. Store: {store_key}.")
    for progress, step, command in build_office_activation_commands(version_label, product_key):
        masked = " ".join("[saved-key]" if part.startswith("/inpkey:") else part for part in command)
        yield _event(progress, step, f"Изпълнява се: {masked}")
        yield _event(progress + 3, step, _run_command(command))
    yield _event(100, "Готово", f"{version_label} activation workflow завърши.", True)


def _saved_office_key(product_label: str) -> tuple[str, str]:
    store = _secure_store()
    store_key = office_store_key(product_label)
    product_key = store.get(store_key, "").strip()
    if product_key:
        return product_key, store_key
    legacy_key = office_legacy_store_key(product_label)
    if legacy_key:
        return store.get(legacy_key, "").strip(), legacy_key
    return "", store_key


def _detect_installed_office() -> tuple[str | None, str, Path | None]:
    ospp = find_ospp_vbs()
    if not ospp:
        return None, "OSPP.VBS не е намерен. Вероятно няма инсталиран Office пакет или инсталацията е нестандартна.", None
    try:
        output = _run_command([locate_cscript_executable(), "//nologo", str(ospp), "/dstatus"])
    except Exception as exc:
        return None, f"OSPP.VBS е намерен, но статусът не може да бъде прочетен:\n{exc}", ospp
    product_label = _office_product_from_status(output)
    return product_label, output, ospp


def _office_product_from_status(output: str) -> str | None:
    text = output.lower()
    if "project" in text:
        prefix = "Project"
    elif "visio" in text:
        prefix = "Visio"
    else:
        prefix = "Office"
    version_match = re.search(r"(2024|2021|2019|2016)", text)
    if version_match:
        return f"{prefix} {version_match.group(1)}"
    if "office24" in text:
        return f"{prefix} 2024"
    if "office21" in text:
        return f"{prefix} 2021"
    if "office19" in text:
        return f"{prefix} 2019"
    if "office16" in text:
        return f"{prefix} 2016"
    if "microsoft 365" in text or "o365" in text or "subscription" in text:
        return "Microsoft 365"
    return None


def _run_auto_office_activation() -> Iterator[dict[str, object]]:
    yield _event(12, "Проверка на Office", "Търси се инсталиран Office пакет и activation script.")
    product_label, status_output, ospp = _detect_installed_office()
    if not ospp:
        yield _event(100, "Office не е открит", status_output, False)
        return
    yield _event(35, "OSPP.VBS е намерен", f"Activation script: {ospp}")
    if not product_label:
        yield _event(100, "Неясна Office версия", "Office е намерен, но версията не може да бъде разпозната автоматично.\n\n" + status_output[:1800], False)
        return
    yield _event(50, "Разпознат Office", f"Намерена версия: {product_label}. Проверява се записан ключ.")
    product_key, store_key = _saved_office_key(product_label)
    if not product_key:
        yield _event(100, "Липсва ключ", f"Няма записан ключ за {office_key_label(product_label)}. Запиши ключ за този Office пакет и опитай отново.", False)
        return
    yield _event(62, "Ключът е намерен", f"Намерен е записан ключ за {product_label}. Store: {store_key}. Ключът няма да се показва в логовете.")
    for progress, step, command in build_office_activation_commands(product_label, product_key):
        adjusted_progress = max(progress, 68)
        masked = " ".join("[saved-key]" if part.startswith("/inpkey:") else part for part in command)
        yield _event(adjusted_progress, step, f"Изпълнява се: {masked}")
        yield _event(min(adjusted_progress + 4, 96), step, _run_command(command))
    yield _event(100, "Готово", f"{product_label} activation workflow завърши.", True)


def _run_office_status() -> Iterator[dict[str, object]]:
    status = check_maintenance_action("office_check_activation_status")
    yield _event(20, "Проверка", status.message)
    if not status.available:
        yield _event(100, "Недостъпно", status.message, False)
        return
    ospp = find_ospp_vbs()
    if not ospp:
        yield _event(100, "Недостъпно", "OSPP.VBS не е намерен.", False)
        return
    output = _run_command([locate_cscript_executable(), "//nologo", str(ospp), "/dstatus"])
    yield _event(100, "Готово", output, True)


def _run_office_repair_check() -> Iterator[dict[str, object]]:
    status = check_maintenance_action("office_quick_repair")
    yield _event(25, "Проверка", status.message)
    if not status.available:
        yield _event(100, "Недостъпно", status.message, False)
        return
    tool = find_click_to_run_executable()
    if not tool:
        yield _event(100, "Недостъпно", "OfficeClickToRun.exe не е намерен.", False)
        return
    command = [
        str(tool),
        "scenario=Repair",
        "platform=x64",
        "culture=en-us",
        "RepairType=FullRepair",
        "DisplayLevel=True",
    ]
    yield _event(55, "Стартиране на Office repair", " ".join(command))
    result = _run_command_result(command)
    output = _command_output(result)
    if result.returncode != 0:
        yield _event(100, "Грешка", output or f"Office repair върна код {result.returncode}.", False)
        return
    yield _event(100, "Готово", output or "Office repair процесът е стартиран успешно.", True)


def _run_office_uninstall_check() -> Iterator[dict[str, object]]:
    status = check_maintenance_action("office_force_uninstall_all")
    yield _event(15, "Проверка", status.message)
    if not status.available:
        yield _event(100, "Недостъпно", status.message, False)
        return
    winget = _winget_executable()
    if not winget:
        yield _event(100, "Winget липсва", "Winget не е намерен.", False)
        return
    output_lines: list[str] = []
    failures: list[str] = []
    total = len(OFFICE_FORCE_UNINSTALL_IDS)
    for index, package_id in enumerate(OFFICE_FORCE_UNINSTALL_IDS, start=1):
        progress = 20 + int((index - 1) * 70 / max(1, total))
        command = [winget, "uninstall", "--id", package_id, "--silent", "--disable-interactivity", "--accept-source-agreements"]
        yield _event(progress, f"Премахване {index}/{total}", " ".join(command))
        result = _run_command_result(command)
        output = _command_output(result)
        output_lines.append(f"[{package_id}]\n{output or 'No output.'}")
        normalized = output.lower()
        if result.returncode != 0 and "no installed package found" not in normalized and "no package found matching input criteria" not in normalized:
            failures.append(package_id)
    final_message = "\n\n".join(output_lines) or "Office cleanup sequence finished."
    if failures:
        final_message += "\n\nFailed packages:\n" + "\n".join(failures)
    yield _event(100, "Готово" if not failures else "Има грешки", final_message, not failures)


def _run_office_offline_install(action_id: str) -> Iterator[dict[str, object]]:
    installer = get_office_offline_installer(action_id)
    missing = []
    if not installer.setup_path.exists():
        missing.append(f"setup.exe липсва: {installer.setup_path}")
    if not installer.config_path.exists():
        missing.append(f"Configuration липсва: {installer.config_path}")
    if missing:
        yield _event(100, "Липсва Office installer", "\n".join(missing), False)
        return
    command = [str(installer.setup_path), "/configure", str(installer.config_path)]
    yield _event(35, "Office installer намерен", f"{installer.label}\n{installer.installers_root}")
    yield _event(70, "Стартиране", " ".join(command))
    result = _run_command_result(command, cwd=installer.setup_path.parent)
    output = _command_output(result)
    if result.returncode != 0:
        yield _event(100, "Грешка", output or f"{installer.label} върна код {result.returncode}.", False)
        return
    yield _event(100, "Готово", output or f"{installer.label} стартира успешно.", True)


def _office_install_architecture() -> str:
    return "64" if os.environ.get("ProgramFiles(x86)") else "32"


def _download_office_deployment_tool(target_dir: Path) -> Path:
    request = urllib.request.Request(ODT_CONFIRMATION_URL, headers={"User-Agent": "Mozilla/5.0 WinSysGuardianAdvanced"})
    with urllib.request.urlopen(request, timeout=30) as response:
        confirmation_html = response.read().decode("utf-8", errors="replace")
    match = re.search(
        r"https://download\.microsoft\.com/download/[^\"'<>\s]+officedeploymenttool_[^\"'<>\s]+\.exe",
        confirmation_html,
        re.IGNORECASE,
    )
    if not match:
        raise RuntimeError("Не намерих валиден Microsoft линк за Office Deployment Tool.")
    odt_url = match.group(0)
    output_path = target_dir / Path(odt_url.split("?")[0]).name
    with urllib.request.urlopen(odt_url, timeout=120) as response, output_path.open("wb") as output_file:
        shutil.copyfileobj(response, output_file)
    return output_path


def _extract_office_deployment_tool(odt_exe: Path, target_dir: Path) -> Path:
    result = _run_command_result([str(odt_exe), "/quiet", f"/extract:{target_dir}"])
    setup_path = target_dir / "setup.exe"
    if setup_path.exists():
        return setup_path
    raise RuntimeError(_command_output(result) or "Office Deployment Tool не се разпакрира правилно.")


def _write_office_online_config(config_path: Path, product_id: str, channel: str) -> None:
    config_text = f"""<Configuration>
  <Add OfficeClientEdition="{_office_install_architecture()}" Channel="{channel}">
    <Product ID="{product_id}">
      <Language ID="MatchOS" Fallback="en-us" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"""
    config_path.write_text(config_text, encoding="utf-8")


def _run_office_online_install(action_id: str) -> Iterator[dict[str, object]]:
    package = get_online_package(action_id)
    status = check_online_package(action_id)
    yield _event(15, "Проверка", status.message)
    if not status.available:
        yield _event(100, "Недостъпно", status.message, False)
        return
    with tempfile.TemporaryDirectory(prefix="wga-office-online-") as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        odt_dir = temp_dir / "odt"
        odt_dir.mkdir(parents=True, exist_ok=True)
        yield _event(30, "Office Deployment Tool", "Изтегля се официалният Office Deployment Tool от Microsoft.")
        odt_exe = _download_office_deployment_tool(odt_dir)
        yield _event(50, "Разархивиране", str(odt_exe))
        setup_exe = _extract_office_deployment_tool(odt_exe, odt_dir)
        config_path = temp_dir / "configuration.xml"
        _write_office_online_config(config_path, package.product_id, package.channel)
        command = [str(setup_exe), "/configure", str(config_path)]
        yield _event(75, "Стартиране", f"Product ID: {package.product_id}\nChannel: {package.channel}")
        result = _run_command_result(command, cwd=odt_dir)
        output = _command_output(result)
        if result.returncode != 0:
            yield _event(100, "Грешка", output or f"{package.label} върна код {result.returncode}.", False)
            return
    yield _event(100, "Готово", output or f"{package.label} стартира успешно чрез Office Deployment Tool.", True)


def _run_nexus_action(action_id: str) -> Iterator[dict[str, object]]:
    status = check_nexus_admin_status()
    yield _event(20, "Проверка", status.message)
    if not status.available:
        yield _event(100, "Недостъпно", status.message, False)
        return
    if action_id == "nexus_list_users":
        result = list_users()
        output = _command_output(result)
        yield _event(100, "Готово" if result.returncode == 0 else "Грешка", output, result.returncode == 0)
        return
    yield _event(
        100,
        "Липсват входни данни",
        "Това Nexus действие не може да се изпълни директно без потребителско име/парола/избор. Стартирай го през новия GUI, който събира нужните данни преди изпълнение.",
        False,
    )


def _run_language_action(action_id: str) -> Iterator[dict[str, object]]:
    yield _event(15, "Проверка", "Чете се текущият Windows language status.")
    status = get_language_status()
    title, command = build_language_action(action_id, status)
    if command == "No command is required.":
        yield _event(100, "Готово", status.summary, True)
        return
    yield _event(55, title, "Прилага се Windows language команда.")
    output = _run_command(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command])
    yield _event(100, "Готово", output, True)


def _run_add_desktop_icons() -> Iterator[dict[str, object]]:
    def progress(value: int, status: str, detail: str) -> None:
        events.append(_event(value, status, detail))

    events: list[dict[str, object]] = []
    yield _event(10, "Подготовка", "Включват се системните икони на работния плот.")
    enabled_labels = enable_windows_desktop_icons(progress)
    for event in events:
        yield event
    yield _event(100, "Готово", "Активирани икони:\n" + "\n".join(f"- {label}" for label in enabled_labels), True)


def _find_onedrive_executable() -> str | None:
    candidates = [
        Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "OneDrive" / "OneDrive.exe",
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Microsoft OneDrive" / "OneDrive.exe",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Microsoft OneDrive" / "OneDrive.exe",
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    return shutil.which("OneDrive.exe")


def _run_onedrive_reset(action_id: str) -> Iterator[dict[str, object]]:
    onedrive_exe = _find_onedrive_executable()
    if not onedrive_exe:
        yield _event(100, "OneDrive липсва", "OneDrive.exe не е намерен на този компютър.", False)
        return
    local_onedrive_dir = str(Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "OneDrive")
    method_steps: dict[str, list[tuple[int, str, str]]] = {
        "reset_onedrive_1": [(35, "Стандартен reset", f'& "{onedrive_exe}" /reset')],
        "reset_onedrive_2": [
            (35, "Спиране на OneDrive", "Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue"),
            (80, "Стартиране на OneDrive", f'Start-Process -FilePath "{onedrive_exe}"'),
        ],
        "reset_onedrive_3": [
            (25, "Спиране на OneDrive", "Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue"),
            (60, "Изчистване на локални OneDrive данни", f'Remove-Item -Recurse -Force "{local_onedrive_dir}" -ErrorAction SilentlyContinue'),
            (90, "Стартиране на OneDrive", f'Start-Process -FilePath "{onedrive_exe}"'),
        ],
    }
    for progress, status, script in method_steps[action_id]:
        yield _event(progress, status, f"PowerShell: {script}")
        output = _run_command(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script])
        if output:
            yield _event(progress, status, output)
    yield _event(100, "Готово", "OneDrive reset операцията завърши.", True)


def _manifest_resource(resource_id: str) -> tuple[dict[str, object] | None, Path, Path]:
    installers_root = resolve_installers_root(_project_root())
    manifest_path = _project_root() / "installers_manifest.json"
    if not manifest_path.exists():
        return None, installers_root, manifest_path
    import json

    manifest = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
    resource = next((item for item in manifest.get("resources", []) if item.get("id") == resource_id), None)
    return resource, installers_root, manifest_path


def _manifest_resources() -> list[dict[str, object]]:
    manifest_path = _project_root() / "installers_manifest.json"
    if not manifest_path.exists():
        return []
    try:
        import json

        manifest = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError):
        return []
    return [item for item in manifest.get("resources", []) if isinstance(item, dict)]


def get_auto_installer_packages() -> list[tuple[str, str]]:
    packages: list[tuple[str, str]] = list(AUTO_INSTALLER_PACKAGES)
    packages.extend((spec["label"], action_id) for action_id, spec in LOCAL_INSTALLER_TASKS.items())
    packages.extend((installer.label, installer.action_id) for installer in OFFICE_OFFLINE_INSTALLERS.values())
    packages.append(("Adobe Reader", "install_adobe_reader"))
    packages.extend((package.label, package.action_id) for package in OFFICE_ONLINE_PACKAGES.values())

    known_resource_ids = {spec["resource_id"] for spec in LOCAL_INSTALLER_TASKS.values()}
    known_resource_ids.add("adobe_reader")
    for resource in _manifest_resources():
        resource_id = str(resource.get("id") or "").strip()
        if not resource_id or resource_id in known_resource_ids or resource_id.startswith("office_"):
            continue
        label = str(resource.get("name") or resource_id).strip()
        packages.append((label, f"resource_{resource_id}"))

    packages.extend(AUTO_INSTALLER_LANGUAGE_PACKAGES)
    packages.extend(AUTO_INSTALLER_POST_INSTALL_ACTIONS)

    deduped: list[tuple[str, str]] = []
    seen: set[str] = set()
    for label, action_id in packages:
        if action_id in seen:
            continue
        seen.add(action_id)
        deduped.append((label, action_id))
    return sorted(deduped, key=lambda item: item[0].casefold())


def get_auto_installer_office_packages() -> list[tuple[str, str]]:
    office_ids = set(OFFICE_OFFLINE_INSTALLERS) | set(OFFICE_ONLINE_PACKAGES)
    return [(label, action_id) for label, action_id in get_auto_installer_packages() if action_id in office_ids]


def _full_package_ids(available_ids: set[str]) -> list[str]:
    office_ids = (set(OFFICE_OFFLINE_INSTALLERS) | set(OFFICE_ONLINE_PACKAGES)) & available_ids
    package_ids = available_ids - office_ids
    preferred_office = next((action_id for action_id in PREFERRED_FULL_PACKAGE_OFFICE_IDS if action_id in office_ids), None)
    if preferred_office:
        package_ids.add(preferred_office)
    return sorted(package_ids)


def get_auto_installer_presets() -> list[dict[str, object]]:
    available_ids = {action_id for _label, action_id in get_auto_installer_packages()}
    presets: list[dict[str, object]] = []
    for preset_id, preset in AUTO_INSTALLER_PRESETS.items():
        raw_packages = preset["packages"]
        package_ids = _full_package_ids(available_ids) if raw_packages == "all" else [package_id for package_id in raw_packages if package_id in available_ids]
        presets.append(
            {
                "id": preset_id,
                "label": preset["label"],
                "description": preset["description"],
                "packages": package_ids,
            }
        )
    return presets


def _resource_required_files(resource_id: str) -> tuple[list[Path], dict[str, object] | None, Path]:
    resource, installers_root, _manifest_path = _manifest_resource(resource_id)
    if not resource:
        return [], None, installers_root
    required_files = [installers_root / str(path) for path in resource.get("required_files", [])]
    return required_files, resource, installers_root


def _run_local_installer_task(action_id: str) -> Iterator[dict[str, object]]:
    spec = LOCAL_INSTALLER_TASKS[action_id]
    resource_id = spec["resource_id"]
    label = spec["label"]
    resource, _installers_root, manifest_path = _manifest_resource(resource_id)
    if not manifest_path.exists():
        yield _event(100, "Липсва manifest", f"Не е намерен {manifest_path}.", False)
        return
    if not resource:
        yield _event(100, "Липсва ресурс", f"В manifest няма ресурс {resource_id}.", False)
        return
    required_files, _resource, _root = _resource_required_files(resource_id)
    missing = [path for path in required_files if not path.exists()]
    if missing:
        yield _event(100, "Липсва installer", "Липсват файлове:\n" + "\n".join(str(path) for path in missing), False)
        return
    detect_value = spec.get("detect_value", "")
    if spec.get("detect_mode") == "winget" and detect_value:
        winget = _winget_executable()
        if winget:
            list_result = _run_command_result([winget, "list", "--id", detect_value, "--exact", "--accept-source-agreements"], timeout=90)
            list_output = _command_output(list_result)
            if list_result.returncode == 0 and detect_value.lower() in list_output.lower():
                yield _event(100, "Пропуснато", f"{label} вече е инсталиран.\n{list_output}", True)
                return
    installer = required_files[0]
    command = [str(installer)]
    silent_args = spec.get("silent_args", "").strip()
    if silent_args:
        command.extend(part for part in silent_args.split(" ") if part)
    yield _event(45, "Installer намерен", f"{label}\n{installer}")
    result = _run_command_result(command, cwd=installer.parent)
    output = _command_output(result)
    if result.returncode != 0:
        yield _event(100, "Грешка", output or f"{label} върна код {result.returncode}.", False)
        return
    yield _event(100, "Готово", output or f"{label} завърши успешно.", True)


def _run_generic_resource_action(action_id: str) -> Iterator[dict[str, object]]:
    resource_id = action_id.removeprefix("resource_")
    resource, _installers_root, manifest_path = _manifest_resource(resource_id)
    if not manifest_path.exists():
        yield _event(100, "Липсва manifest", f"Не е намерен {manifest_path}.", False)
        return
    if not resource:
        yield _event(100, "Липсва ресурс", f"В manifest няма ресурс {resource_id}.", False)
        return
    required_files, _resource, _root = _resource_required_files(resource_id)
    missing = [path for path in required_files if not path.exists()]
    if missing:
        yield _event(100, "Липсва ресурс", "Липсват файлове:\n" + "\n".join(str(path) for path in missing), False)
        return
    local_file = required_files[0]
    label = str(resource.get("name") or resource_id)
    if local_file.suffix.lower() not in {".exe", ".msi", ".bat", ".cmd"}:
        yield _event(100, "Наличен ресурс", f"{label}: {local_file}", True)
        return
    command = ["msiexec", "/i", str(local_file)] if local_file.suffix.lower() == ".msi" else ["cmd", "/c", str(local_file)] if local_file.suffix.lower() in {".bat", ".cmd"} else [str(local_file)]
    yield _event(45, "Стартиране на ресурс", f"{label}\n{local_file}")
    result = _run_command_result(command, cwd=local_file.parent)
    output = _command_output(result)
    if result.returncode != 0:
        yield _event(100, "Грешка", output or f"{label} върна код {result.returncode}.", False)
        return
    yield _event(100, "Готово", output or f"{label} стартира успешно.", True)


def _run_local_installer_task_once(action_id: str) -> tuple[bool, str]:
    spec = LOCAL_INSTALLER_TASKS[action_id]
    resource_id = spec["resource_id"]
    label = spec["label"]
    resource, _installers_root, manifest_path = _manifest_resource(resource_id)
    if not manifest_path.exists():
        return False, f"Не е намерен {manifest_path}."
    if not resource:
        return False, f"В manifest няма ресурс {resource_id}."
    required_files, _resource, _root = _resource_required_files(resource_id)
    missing = [path for path in required_files if not path.exists()]
    if missing:
        return False, "Липсват файлове:\n" + "\n".join(str(path) for path in missing)
    detect_value = spec.get("detect_value", "")
    if spec.get("detect_mode") == "winget" and detect_value:
        winget = _winget_executable()
        if winget:
            list_result = _run_command_result([winget, "list", "--id", detect_value, "--exact", "--accept-source-agreements"], timeout=90)
            list_output = _command_output(list_result)
            if list_result.returncode == 0 and detect_value.lower() in list_output.lower():
                return True, f"Пропуснато: {label} вече е инсталиран.\n{list_output}"
    installer = required_files[0]
    command = [str(installer)]
    silent_args = spec.get("silent_args", "").strip()
    if silent_args:
        command.extend(part for part in silent_args.split(" ") if part)
    result = _run_command_result(command, cwd=installer.parent)
    output = _command_output(result)
    return result.returncode == 0, output or f"{label}: return code {result.returncode}"


def _run_generic_resource_once(action_id: str) -> tuple[bool, str]:
    resource_id = action_id.removeprefix("resource_")
    resource, _installers_root, manifest_path = _manifest_resource(resource_id)
    if not manifest_path.exists():
        return False, f"Не е намерен {manifest_path}."
    if not resource:
        return False, f"В manifest няма ресурс {resource_id}."
    required_files, _resource, _root = _resource_required_files(resource_id)
    missing = [path for path in required_files if not path.exists()]
    if missing:
        return False, "Липсват файлове:\n" + "\n".join(str(path) for path in missing)
    local_file = required_files[0]
    label = str(resource.get("name") or resource_id)
    if local_file.suffix.lower() not in {".exe", ".msi", ".bat", ".cmd"}:
        return True, f"Наличен ресурс: {local_file}"
    command = ["msiexec", "/i", str(local_file)] if local_file.suffix.lower() == ".msi" else ["cmd", "/c", str(local_file)] if local_file.suffix.lower() in {".bat", ".cmd"} else [str(local_file)]
    result = _run_command_result(command, cwd=local_file.parent)
    output = _command_output(result)
    return result.returncode == 0, output or f"{label}: return code {result.returncode}"


def _run_resources_status() -> Iterator[dict[str, object]]:
    yield _event(12, "Проверка", "Проверяват се инсталационните ресурси от installers_manifest.json.")
    status = check_resource_status(_project_root())
    report = missing_resource_report(status)
    detail = (
        f"Installers: {status.installers_root}\n"
        f"Налични: {status.available}/{status.total}\n"
        f"Липсват: {status.missing}\n"
        f"За изтегляне: {status.downloadable_missing}\n\n"
        f"{report}"
    )
    yield _event(100, "Готово" if status.complete else "Има липсващи ресурси", detail, status.configured)


def _run_resources_download_missing() -> Iterator[dict[str, object]]:
    status = check_resource_status(_project_root())
    yield _event(8, "Проверка", missing_resource_report(status))
    missing_downloads = [check.item for check in status.checks if not check.available and check.item.url]
    if not status.configured:
        yield _event(100, "Липсва manifest", "Няма конфигуриран manifest за инсталационните ресурси.", False)
        return
    if not missing_downloads:
        final_status = check_resource_status(_project_root())
        yield _event(100, "Няма ресурси за изтегляне", missing_resource_report(final_status), final_status.complete)
        return
    total_items = len(missing_downloads)
    errors: list[str] = []
    for index, item in enumerate(missing_downloads, start=1):
        base_progress = int(((index - 1) / total_items) * 90) + 5
        yield _event(base_progress, f"Изтегляне {index}/{total_items}", item.name)

        def progress(downloaded: int, total: int, name: str, speed: float = 0.0, eta: int = 0, phase: str = "") -> None:
            percent = int((downloaded / total) * 100) if total else 0
            total_progress = int(((index - 1) + (percent / 100)) * 90 / total_items) + 5
            action = phase or "Изтегляне"
            progress_events.put(
                _event(
                    min(96, total_progress),
                    action,
                    f"{name}: {percent}% | {downloaded / 1048576:.1f} / {total / 1048576:.1f} MB | Скорост: {speed / 1048576:.2f} MB/s | ETA: {eta}s",
                )
            )

        progress_events = Queue()

        def transfer():
            try:
                progress_events.put((True, download_resource(_project_root(), item, progress)))
            except Exception as exc:
                progress_events.put((False, exc))

        try:
            yield _event(base_progress, "Източник", f"{item.name}\n{item.url}")
            thread = Thread(target=transfer, daemon=True)
            thread.start()
            while True:
                event = progress_events.get()
                if isinstance(event, tuple):
                    thread.join()
                    if not event[0]:
                        raise event[1]
                    target_path = event[1]
                    break
                yield event
            yield _event(min(96, int((index / total_items) * 90) + 5), "Готов пакет", f"{item.name}: {target_path}")
        except Exception as exc:
            errors.append(f"{item.name}: {exc}")
            yield _event(min(96, int((index / total_items) * 90) + 5), "Проблем", f"{item.name}: {exc}")
    final_status = check_resource_status(_project_root())
    if errors:
        yield _event(100, "Има грешки", "\n".join(errors) + "\n\n" + missing_resource_report(final_status), False)
        return
    yield _event(100, "Готово", "Липсващите инсталационни ресурси са изтеглени.\n\n" + missing_resource_report(final_status), final_status.complete)


def _consume_action_events(events: Iterator[dict[str, object]]) -> tuple[bool, str]:
    lines: list[str] = []
    success = False
    for event in events:
        status = str(event.get("status") or "")
        detail = str(event.get("detail") or "")
        if detail:
            lines.append(f"{status}: {detail}")
        if bool(event.get("finished")):
            success = True
    return success, "\n".join(lines[-8:]) or "Операцията приключи."


def _run_adobe_reader_install() -> Iterator[dict[str, object]]:
    status = check_adobe_reader_status(_project_root())
    yield _event(20, "Проверка", status.message)
    winget_exe = _winget_executable()
    if not winget_exe:
        yield _event(100, "Winget липсва", "Winget не е намерен. Adobe Reader не може да бъде инсталиран онлайн.", False)
        return
    command = [
        winget_exe,
        "install",
        "--id",
        ADOBE_READER_WINGET_ID,
        "--exact",
        "--source",
        "winget",
        "--silent",
        "--disable-interactivity",
        "--accept-package-agreements",
        "--accept-source-agreements",
    ]
    yield _event(55, "Инсталиране", " ".join(command))
    result = _run_command_result(command, timeout=900)
    output = _command_output(result)
    if result.returncode != 0:
        yield _event(100, "Грешка", output or f"Adobe Reader installer върна код {result.returncode}.", False)
        return
    yield _event(100, "Готово", output or "Adobe Reader беше инсталиран/обновен успешно.", True)


def _winget_executable() -> str | None:
    from office_online import find_winget_executable

    return find_winget_executable()


def _run_winget_install(package_id: str) -> tuple[bool, str]:
    winget = _winget_executable()
    if not winget:
        return False, "Winget не е намерен."
    list_result = _run_command_result([winget, "list", "--id", package_id, "--exact", "--accept-source-agreements"], timeout=90)
    list_output = _command_output(list_result)
    if list_result.returncode == 0 and package_id.lower() in list_output.lower():
        return True, f"Пропуснато: {package_id} вече е инсталиран.\n{list_output}"
    command = [
        winget,
        "install",
        "--id",
        package_id,
        "--exact",
        "--source",
        "winget",
        "--silent",
        "--disable-interactivity",
        "--accept-package-agreements",
        "--accept-source-agreements",
    ]
    result = _run_command_result(command, timeout=1800)
    output = _command_output(result)
    return result.returncode == 0, output or f"{package_id}: return code {result.returncode}"


def _run_secret_install_action(action_id: str) -> Iterator[dict[str, object]]:
    winget = _winget_executable()
    if not winget:
        yield _event(100, "Winget липсва", "Winget не е намерен. Secret Install preset-ите изискват winget.", False)
        return
    if action_id == "update_all_apps":
        command = [winget, "upgrade", "--all", "--silent", "--disable-interactivity", "--accept-package-agreements", "--accept-source-agreements"]
        yield _event(35, "Актуализация", " ".join(command))
        result = _run_command_result(command, timeout=3600)
        output = _command_output(result)
        yield _event(100, "Готово" if result.returncode == 0 else "Грешка", output or f"winget upgrade върна код {result.returncode}.", result.returncode == 0)
        return
    packages = SECRET_INSTALL_PRESETS[action_id]
    completed = 0
    failed: list[str] = []
    output_lines: list[str] = []
    total = len(packages)
    for index, package_id in enumerate(packages, start=1):
        progress = int(((index - 1) / total) * 90) + 5
        yield _event(progress, f"Инсталиране {index}/{total}", package_id)
        success, output = _run_winget_install(package_id)
        completed += 1 if success else 0
        if not success:
            failed.append(package_id)
        output_lines.append(f"[{package_id}]\n{output}")
    summary = f"Завършени: {completed}/{total}\nНеуспешни: {len(failed)}"
    if failed:
        summary += "\nFailed: " + ", ".join(failed)
    yield _event(100, "Готово" if not failed else "Има грешки", summary + "\n\n" + "\n\n".join(output_lines[-3:]), not failed)


def _run_driver_action(action_id: str) -> Iterator[dict[str, object]]:
    if action_id == "driver_backup_advanced":
        yield _event(100, "Отвори Driver Tools", "Разширеният driver workflow е разделен на реални действия: clean backup, full backup, recovery USB, restore и PC report.", True)
        return
    base = _driver_tools_base()
    if action_id == "driver_pc_report":
        yield _event(35, "Събиране на данни", "Генерира се PC report.")
        report = generate_pc_report(base)
        yield _event(100, "Готово", f"PC report е създаден: {report}", True)
        return
    if action_id in {"driver_recovery_usb", "driver_restore_last"}:
        backups = [path for path in base.glob("DriverBackup_*") if path.is_dir()]
        if not backups:
            yield _event(100, "Липсва backup", f"Няма намерен driver backup в {base}. Първо направи clean/full backup.", False)
            return
        latest_backup = max(backups, key=lambda path: path.stat().st_mtime)
        if action_id == "driver_restore_last":
            yield _event(35, "Възстановяване", f"Използва се последният backup: {latest_backup}")
            result = restore_drivers_from_backup(latest_backup)
            output = _command_output(result)
            if result.returncode != 0:
                yield _event(100, "Грешка", output or "Driver restore failed.", False)
                return
            yield _event(100, "Готово", output or f"Драйверите са възстановени от {latest_backup}.", True)
            return
        drives = detect_removable_drives()
        if not drives:
            yield _event(100, "Няма USB", "Не е открит removable USB drive. Свържи USB устройство и опитай отново.", False)
            return
        usb_root = drives[0]
        yield _event(45, "Копиране към USB", f"Backup: {latest_backup}\nUSB: {usb_root}")
        recovery_dir, restore_script = create_recovery_usb(latest_backup, usb_root)
        yield _event(100, "Готово", f"Recovery folder: {recovery_dir}\nRestore script: {restore_script}", True)
        return
    mode = "clean" if action_id == "driver_backup_clean" else "full"
    destination = create_backup_folder(base)
    yield _event(20, "Подготовка", f"Backup папка: {destination}")
    result, log_path = export_drivers(destination, mode)
    if result.returncode != 0:
        raise RuntimeError((result.stderr or result.stdout or "Driver export failed.").strip())
    yield _event(70, "Списък с драйвери", str(create_driver_list(destination)))
    restore_script = create_restore_script(destination)
    archive = compress_backup(destination)
    yield _event(100, "Готово", f"Backup: {destination}\nZIP: {archive}\nLog: {log_path}\nRestore: {restore_script}", True)


def _driver_tools_base() -> Path:
    candidates = [desktop_path() / "WGA Driver Tools", _project_root() / "WGA Driver Tools"]
    last_error: OSError | None = None
    for candidate in candidates:
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            return candidate
        except OSError as exc:
            last_error = exc
    raise RuntimeError(f"Не може да се създаде driver tools папка: {last_error}")


def _run_auto_installer(selected_packages: list[str] | None = None) -> Iterator[dict[str, object]]:
    selected = set(selected_packages or [])
    packages = [(label, package_id) for label, package_id in get_auto_installer_packages() if not selected or package_id in selected]
    if not packages:
        yield _event(100, "Няма избор", "Не е избрано нито едно приложение за инсталиране.", False)
        return
    total = len(packages)
    completed = 0
    failed: list[str] = []
    skipped: list[str] = []
    output_lines: list[str] = []
    for index, (label, package_id) in enumerate(packages, start=1):
        progress = int(((index - 1) / total) * 95) + 3
        yield _event(progress, f"Application {index}/{total}", f"{label}: Waiting/Checking")
        if package_id in LOCAL_INSTALLER_TASKS:
            success, output = _run_local_installer_task_once(package_id)
        elif package_id.startswith("resource_"):
            success, output = _run_generic_resource_once(package_id)
        elif package_id == "install_adobe_reader":
            success, output = _consume_action_events(_run_adobe_reader_install())
        elif package_id in OFFICE_OFFLINE_INSTALLERS:
            success, output = _consume_action_events(_run_office_offline_install(package_id))
        elif package_id.startswith("online_"):
            success, output = _consume_action_events(_run_office_online_install(package_id))
        elif package_id in {action_id for _label, action_id in AUTO_INSTALLER_LANGUAGE_PACKAGES}:
            success, output = _consume_action_events(_run_language_action(package_id))
        elif package_id == "auto_windows_activation":
            success, output = _consume_action_events(_run_auto_windows_activation())
        elif package_id == "auto_office_activation":
            success, output = _consume_action_events(_run_auto_office_activation())
        else:
            success, output = _run_winget_install(package_id)
        if "Пропуснато:" in output:
            skipped.append(label)
        elif success:
            completed += 1
        else:
            failed.append(label)
        output_lines.append(f"[{label}]\n{output}")
        next_progress = int((index / total) * 95)
        status = "Completed" if success else "Failed"
        yield _event(next_progress, status, f"{label}: {status}")
    summary = (
        f"Completed: {completed}\n"
        f"Skipped: {len(skipped)}\n"
        f"Failed: {len(failed)}\n"
        f"Remaining: 0"
    )
    if skipped:
        summary += "\nSkipped apps: " + ", ".join(skipped)
    if failed:
        summary += "\nFailed apps: " + ", ".join(failed)
    yield _event(100, "Готово" if not failed else "Има грешки", summary + "\n\n" + "\n\n".join(output_lines[-4:]), not failed)
