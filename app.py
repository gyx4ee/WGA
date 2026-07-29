# Основният файл събира прозорците, менюто, dashboard екрана и връзките към помощните модули.
from __future__ import annotations

import base64
import ctypes
import hashlib
import math
import os
import platform
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import tkinter.font as tkfont
import traceback
import json
import webbrowser
import winreg
import urllib.error
import urllib.request
from tkinter import messagebox, simpledialog, ttk
from datetime import datetime
from pathlib import Path

from adobe_reader import ADOBE_READER_WINGET_ID, check_adobe_reader_status
from driver_backup import (
    create_backup_folder,
    create_driver_list,
    create_recovery_usb,
    create_restore_script,
    desktop_path,
    detect_removable_drives,
    export_drivers,
    generate_pc_report,
    onedrive_path,
    compress_backup,
    restore_drivers_from_backup,
)
from office_activation import build_office_activation_commands, get_office_version_label
from office_inventory import detect_installed_office
from office_installers import OFFICE_OFFLINE_INSTALLERS, get_office_offline_installer
from language_manager import build_language_action, get_language_status
from language_manager import LanguageStatus
from nexus_admin import (
    change_password,
    check_nexus_admin_status,
    create_user,
    delete_user,
    list_users,
    set_admin_rights,
    user_details,
)
from office_maintenance import (
    OFFICE_FORCE_UNINSTALL_IDS,
    check_maintenance_action,
    find_click_to_run_executable,
    find_ospp_vbs,
)
from office_online import (
    OFFICE_ONLINE_PACKAGES,
    ODT_CONFIRMATION_URL,
    check_online_package,
    find_winget_executable,
    get_online_package,
)
from path_utils import get_runtime_storage_info
from resource_manager import (
    ResourceStatus,
    check_resource_status,
    download_resource,
    load_resource_manifest,
    missing_resource_report,
)
from self_updater import launch_helper_and_exit, prepare_update_install
from system_health import HealthItem, collect_health_items
from update_checker import UpdateResult, check_for_updates
from optimization_ui import OptimizationUI
from network_monitoring_ui import NetworkMonitoringUI


APP_TITLE = "WinSys Guardian Advanced"
WINDOW_SIZE = "930x630"
MAIN_WINDOW_SIZE = "1220x820"
MAIN_MIN_WIDTH = 1120
MAIN_MIN_HEIGHT = 760
BASE_DPI = 96.0
SPLASH_BASE_WIDTH = 930
SPLASH_BASE_HEIGHT = 630
MAIN_CARD_COLUMNS = 3
PROJECT_ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
RESOURCE_ROOT = Path(getattr(sys, "_MEIPASS", PROJECT_ROOT)).resolve()
SETTINGS_FILE = PROJECT_ROOT / "settings.json"
SECURE_STORE_FILE = PROJECT_ROOT / ".wga_secure_store.json"
AGENT_STATUS_FILE = PROJECT_ROOT / "wga_agent_status.json"
DASHBOARD_ICON_SHEET_RELATIVE = "assets/dashboard-icon-sheet.png"
DASHBOARD_ICONS_MANIFEST_RELATIVE = "assets/dashboard-icons/dashboard-icons.json"
APP_LOGO_RELATIVE = "assets/wga-icon.png"
MENU_ICONS_MANIFEST_RELATIVE = "assets/menu-icons/menu-icons.json"

APP_BG = "#071311"
APP_PANEL = "#0d1c1a"
APP_PANEL_ALT = "#102725"
APP_PANEL_SOFT = "#122f2a"
APP_BORDER = "#1f4e46"
APP_BORDER_STRONG = "#2a6f60"
APP_TEXT = "#ecfff7"
APP_TEXT_SOFT = "#a6d5c5"
APP_TEXT_MUTED = "#7ca394"
APP_ACCENT = "#37e39a"
APP_ACCENT_SOFT = "#1d8f67"
APP_ACCENT_BLUE = "#2f8fff"
APP_WARNING = "#d0a94a"
APP_DANGER = "#c94d58"

SIDEBAR_SECTIONS: tuple[tuple[str, str], ...] = (
    ("main", "Обзор"),
    ("activation", "Активация"),
    ("install_software", "Софтуер"),
    ("language", "Езици"),
    ("auto_installer", "Авто инсталатор"),
    ("driver_backup", "Архивиране"),
    ("nexus_admin", "Nexus Admin"),
)


# Връща път към файл спрямо runtime папката на приложението.
def runtime_file(relative_path: str) -> Path:
    # Търси файла на правилното място според това дали работим от build или от проект.
    portable_path = PROJECT_ROOT / relative_path
    bundled_path = RESOURCE_ROOT / relative_path

    # При build първо ползваме bundled файла, за да не четем старо копие до exe-то.
    if getattr(sys, "frozen", False) and bundled_path.exists():
        return bundled_path

    # При portable или dev режим ползваме файла до програмата, ако го има.
    if portable_path.exists():
        return portable_path

    # Ако няма локално копие, връщаме bundled ресурса като резервен вариант.
    return bundled_path


# Ограничава числова стойност в зададени граници.
def clamp(value: float, minimum: float, maximum: float) -> float:
    # Ograni4ava stoinost v bezopasen diapazon.
    return max(minimum, min(maximum, value))


# Настройва DPI поведението на приложението под Windows.
def configure_windows_dpi_awareness() -> None:
    # Kazva na Windows, che prilojenieto trqbva da se ma6abira pravilno na razlichni monitori.
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


# Синхронизира мащабирането на tkinter с текущия DPI.
def apply_tk_dpi_scaling(window: tk.Misc) -> float:
    # Nastroyva Tkinter spored realniq DPI, za da ne izliza droben ili zamazan tekst.
    try:
        dpi = float(window.winfo_fpixels("1i"))
    except Exception:
        dpi = BASE_DPI
    tk_scaling = clamp(dpi / 72.0, 1.0, 2.4)
    try:
        window.tk.call("tk", "scaling", tk_scaling)
    except tk.TclError:
        pass
    font_scale = clamp(dpi / BASE_DPI, 0.95, 1.35)
    for font_name in ("TkDefaultFont", "TkTextFont", "TkHeadingFont", "TkMenuFont", "TkCaptionFont", "TkSmallCaptionFont", "TkIconFont", "TkTooltipFont"):
        try:
            named_font = tkfont.nametofont(font_name)
            base_size = abs(int(named_font.cget("size")))
            if base_size:
                named_font.configure(size=max(9, int(base_size * font_scale)))
        except Exception:
            continue
    return dpi


# Изчислява размер на прозорец според екрана.
def responsive_window_size(screen_width: int, screen_height: int, base_width: int, base_height: int) -> tuple[int, int]:
    # Smalqva ili uvelichava prozoreca според monitora, za da se vizhda qsno bez da izliza ot ekrana.
    width_scale = (screen_width - 80) / max(1, base_width)
    height_scale = (screen_height - 110) / max(1, base_height)
    scale = clamp(min(width_scale, height_scale), 0.78, 1.35)
    return max(780, int(base_width * scale)), max(520, int(base_height * scale))


# Помощна функция за center geometry.
def center_geometry(window: tk.Misc, width: int, height: int) -> None:
    # Centrira prozoreca na ekrana.
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    position_x = max(0, (screen_width - width) // 2)
    position_y = max(0, (screen_height - height) // 2)
    window.geometry(f"{width}x{height}+{position_x}+{position_y}")


# Помощна функция за apply main window layout.
def apply_main_window_layout(window: tk.Tk) -> None:
    # Podbira podhodyasht razmer za glavniq ekran spored rezolyuciqta na monitora.
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    target_width = max(980, screen_width - max(40, screen_width // 25))
    target_height = max(700, screen_height - max(70, screen_height // 14))
    target_width = min(target_width, screen_width - 20)
    target_height = min(target_height, screen_height - 20)
    center_geometry(window, target_width, target_height)
    window.minsize(min(max(900, target_width - 180), target_width), min(max(640, target_height - 120), target_height))
    if screen_width >= 1500 and screen_height >= 850:
        try:
            window.state("zoomed")
        except tk.TclError:
            pass


VERSION_FILE = runtime_file("version.json")
APP_ICON_FILE = runtime_file("assets/wga-icon.ico")
DEFAULT_WINDOWS11_MENU_PASSWORD = "Zinzibar2"
CARD_COLUMNS = 2
CARDS_PER_PAGE = 6
CARD_BUTTON_WIDTH = 26
CARD_BUTTON_HEIGHT = 2
CARD_BUTTON_PIXEL_WIDTH = 260
CARD_BUTTON_PIXEL_HEIGHT = 48
CARD_ACTION_HEIGHT = 52
CARD_ACTION_DOUBLE_HEIGHT = 108
NAV_BUTTON_WIDTH = 11
CARD_MIN_HEIGHT = 185
MENU_CARD_MIN_HEIGHT = {
    "office_center": 215,
    "nexus_admin": 215,
    "language": 185,
    "office_install_center": 215,
    "secret_install": 215,
    "driver_backup": 215,
}
DESKTOP_ICON_PATHS = (
    r"Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel",
    r"Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu",
)
DESKTOP_ICON_TARGETS = (
    ("This PC", "{20D04FE0-3AEA-1069-A2D8-08002B30309D}"),
    ("Network", "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}"),
    ("Control Panel", "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}"),
    ("User Files", "{59031a47-3f72-44a7-89c5-5595fe6b30ee}"),
)
MENU_PAGE_SIZE: dict[str, int] = {
    "activation": 4,
    "reset_onedrive": 4,
    "windows11_activation": 4,
    "office_activation": 4,
    "install_software": 4,
    "office_install_center": 2,
    "auto_installer": 1,
    "secret_install": 4,
    "office_center": 2,
    "language": 4,
    "driver_backup": 2,
    "nexus_admin": 4,
}

OFFICE_ACTION_IDS = [
    "office_2016_activation",
    "office_2019_activation",
    "office_2021_activation",
]

PROGRAM_SELECTOR_LOCAL_TASKS: tuple[dict[str, str], ...] = (
    {
        "id": "install_ninite",
        "label": "Ninite",
        "category": "Основен софтуер",
        "description": "Стартира локалния Ninite installer от папката Installers.",
        "type": "local_installer",
        "resource_id": "ninite_installer",
        "detect_mode": "none",
        "silent_args": "",
    },
    {
        "id": "install_visual_studio_setup",
        "label": "Visual Studio Setup",
        "category": "Development",
        "description": "Стартира локалния installer за Visual Studio от Installers папката.",
        "type": "local_installer",
        "resource_id": "visual_studio_setup",
        "detect_mode": "winget",
        "detect_value": "Microsoft.VisualStudio.2022.Community",
        "silent_args": "",
    },
    {
        "id": "install_vscode_arm64",
        "label": "VS Code ARM64",
        "category": "Development",
        "description": "Инсталира локалния VS Code installer, ако е наличен.",
        "type": "local_installer",
        "resource_id": "vscode_user_setup_arm64",
        "detect_mode": "winget",
        "detect_value": "Microsoft.VisualStudioCode",
        "silent_args": "/VERYSILENT /NORESTART",
    },
)


MENU_TREE = {
    "main": {
        "title": "Main Menu",
        "subtitle": "Central control hub for deployment, activation, language, recovery and admin tools.",
        "items": [
            {"label": "Hidden Agent Menu", "kind": "menu", "target": "hidden_menu", "description": "Direct access to the hidden agent status menu."},
            {"label": "Activation Menu", "kind": "menu", "target": "activation"},
            {
                "label": "Add Desktop Icons",
                "kind": "action",
                "action_id": "add_desktop_icons",
                "description": "Create standard support shortcuts.",
            },
            {
                "label": "Reset OneDrive",
                "kind": "menu",
                "target": "reset_onedrive",
            },
            {"label": "Install Software", "kind": "menu", "target": "install_software"},
            {
                "label": "Автоматичен инсталатор",
                "kind": "menu",
                "target": "auto_installer",
                "description": "Избери няколко инсталации и ги стартирай с едно копче.",
            },
            {"label": "Language Menu", "kind": "menu", "target": "language"},
            {
                "label": "Архивиране",
                "kind": "menu",
                "target": "driver_backup",
                "description": "Архив на драйвери, recovery носител и отчет за компютъра.",
            },
            {
                "label": "System Commander: Nexus Admin",
                "kind": "menu",
                "target": "nexus_admin",
                "description": "Local user management and administrator account tools.",
            },
            {"label": "Hidden Agent Menu", "kind": "menu", "target": "hidden_menu", "description": "Direct access to the hidden agent status menu."},
            {"label": "Reset Console", "kind": "action", "description": "Refresh the current interface state."},
            {"label": "Exit", "kind": "exit", "description": "Close WinSys Guardian Advanced."},
        ],
    },
    "auto_installer": {
        "title": "Автоматичен инсталатор",
        "subtitle": "Отметни нужните програми и настройки, след това ги инсталирай наведнъж.",
        "items": [],
    },
    "activation": {
        "title": "Activation Menu",
        "subtitle": "Windows and Office activation shortcuts.",
        "items": [
            {"label": "Activate Windows 10", "kind": "menu", "target": "windows10_activation", "icon": "key", "accent": "#2f8fff"},
            {"label": "Activate Windows 11", "kind": "menu", "target": "windows11_activation", "icon": "shield", "accent": "#37e39a"},
            {"label": "Office Activation Center", "kind": "menu", "target": "office_activation", "icon": "actions", "accent": "#d0a94a"},
        ],
    },
    "windows10_activation": {
        "title": "Windows 10 Key Manager",
        "subtitle": "Save and manage the Windows 10 product key used by your admin team.",
        "items": [
            {
                "label": "Run Windows 10 Activation",
                "kind": "action",
                "action_id": "activate_windows10",
                "icon": "key",
                "description": "Run the Windows activation commands using the saved product key.",
            },
            {
                "label": "Save or Replace Product Key",
                "kind": "action",
                "action_id": "save_windows10_key",
                "icon": "download",
                "description": "Store the currently approved Windows 10 key for later use.",
            },
            {
                "label": "Show Saved Product Key",
                "kind": "action",
                "action_id": "show_windows10_key",
                "icon": "monitor",
                "description": "Display the key currently stored in this application.",
            },
            {
                "label": "Clear Saved Product Key",
                "kind": "action",
                "action_id": "clear_windows10_key",
                "icon": "warning",
                "description": "Remove the saved key if your organization replaces it.",
            },
            {
                "label": "Return to Main Menu",
                "kind": "menu",
                "target": "main",
            },
        ],
    },
    "windows11_activation": {
        "title": "Windows 11 Key Manager",
        "subtitle": "Save and manage the Windows 11 product key used by your admin team.",
        "items": [
            {
                "label": "Run Windows 11 Activation",
                "kind": "action",
                "action_id": "activate_windows11",
                "icon": "key",
                "description": "Run the Windows activation commands using the saved product key.",
            },
            {
                "label": "Save or Replace Product Key",
                "kind": "action",
                "action_id": "save_windows11_key",
                "icon": "download",
                "description": "Store the currently approved Windows 11 key for later use.",
            },
            {
                "label": "Show Saved Product Key",
                "kind": "action",
                "action_id": "show_windows11_key",
                "icon": "monitor",
                "description": "Display the key currently stored in this application.",
            },
            {
                "label": "Clear Saved Product Key",
                "kind": "action",
                "action_id": "clear_windows11_key",
                "icon": "warning",
                "description": "Remove the saved key if your organization replaces it.",
            },
            {
                "label": "Return to Main Menu",
                "kind": "menu",
                "target": "main",
            },
        ],
    },
    "office_activation": {
        "title": "Office Activation Center",
        "subtitle": "Save and manage the Office product key, then choose the target Office version.",
        "items": [
            {
                "label": "Save or Replace Office Key",
                "kind": "action",
                "action_id": "save_office_key",
                "icon": "download",
                "description": "Store the Office product key used by your admin workflow.",
            },
            {
                "label": "Show Saved Office Key",
                "kind": "action",
                "action_id": "show_office_key",
                "icon": "monitor",
                "description": "Display the Office key currently stored in this application.",
            },
            {
                "label": "Clear Saved Office Key",
                "kind": "action",
                "action_id": "clear_office_key",
                "icon": "warning",
                "description": "Remove the saved Office key if your organization replaces it.",
            },
            {
                "label": "Office 2016",
                "kind": "action",
                "action_id": "office_2016_activation",
                "icon": "key",
                "description": "Run activation workflow for Office 2016 using the saved Office key.",
            },
            {
                "label": "Office 2019",
                "kind": "action",
                "action_id": "office_2019_activation",
                "icon": "key",
                "description": "Run activation workflow for Office 2019 using the saved Office key.",
            },
            {
                "label": "Office 2021",
                "kind": "action",
                "action_id": "office_2021_activation",
                "icon": "key",
                "description": "Run activation workflow for Office 2021 using the saved Office key.",
            },
            {
                "label": "Return to Main Menu",
                "kind": "menu",
                "target": "main",
            },
        ],
    },
    "reset_onedrive": {
        "title": "Reset OneDrive",
        "subtitle": "Choose a reset workflow for the OneDrive client.",
        "items": [
            {
                "label": "Reset OneDrive (Method 1)",
                "kind": "action",
                "action_id": "reset_onedrive_1",
                "description": "Стандартен reset на OneDrive. Подходящ при блокирал sync или липсваща икона.",
            },
            {
                "label": "Reset OneDrive (Method 2)",
                "kind": "action",
                "action_id": "reset_onedrive_2",
                "description": "Спира процеса и стартира OneDrive отново. Полезно при забил процес или замръзнал клиент.",
            },
            {
                "label": "Reset OneDrive (Method 3)",
                "kind": "action",
                "action_id": "reset_onedrive_3",
                "description": "Изтрива локалните OneDrive файлове в профила и прави чисто стартиране. Използвай само ако другите методи не помогнат.",
            },
            {"label": "Return to Main Menu", "kind": "menu", "target": "main"},
        ],
    },
    "hidden_menu": {
        "title": "Hidden Menu",
        "subtitle": "Скрито меню за бърз достъп към специални действия.",
        "items": [
            {
                "label": "Show hidden status",
                "kind": "action",
                "action_id": "hidden_show_status",
                "description": "Показва текущия статус на приложението в прозорец.",
            },
            {
                "label": "Load agent status",
                "kind": "action",
                "action_id": "hidden_load_agent_status",
                "description": "Чете локалния agent статусен файл и показва информация за машината.",
            },
            {"label": "Return to Main Menu", "kind": "menu", "target": "main"},
        ],
    },
    "install_software": {
        "title": "Install Software",
        "subtitle": "Office installers, app deployment and advanced install hubs.",
        "items": [
            {
                "label": "Office Install Center",
                "kind": "menu",
                "target": "office_install_center",
                "description": "Всички Office offline и online инсталации са събрани тук в едно меню.",
            },
            {"label": "Install Ninite", "kind": "action", "action_id": "install_ninite"},
            {
                "label": "Install Adobe Reader",
                "kind": "action",
                "action_id": "install_adobe_reader",
                "description": "Проверява актуалната Adobe Reader версия през winget и предупреждава, ако локалният installer е стар.",
            },
            {"label": "Secret Install Interface", "kind": "menu", "target": "secret_install"},
            {"label": "Return to Main Menu", "kind": "menu", "target": "main"},
        ],
    },
    "office_install_center": {
        "title": "Office Install Center",
        "subtitle": "Обединен център за Office offline и online инсталации от Installers папката.",
        "items": [
            {
                "label": "Office 2016 Offline",
                "kind": "action",
                "action_id": "install_office_2016_offline",
                "description": "Стартира local setup.exe с Configuration.xml от G:\\Installers\\Office2016.",
            },
            {
                "label": "Office 2019 Offline",
                "kind": "action",
                "action_id": "install_office_2019_offline",
                "description": "Стартира local setup.exe с Configuration.xml от G:\\Installers\\Office2019.",
            },
            {
                "label": "Office 2021 Offline",
                "kind": "action",
                "action_id": "install_office_2021_offline",
                "description": "Стартира local setup.exe с Configuration.xml от G:\\Installers\\Office2021.",
            },
            {
                "label": "Office Professional 2021 Offline",
                "kind": "action",
                "action_id": "install_office_2021_new_offline",
                "description": "Опитва инсталация от G:\\Installers\\Office prof 2021, ако файловете са налични.",
            },
            {
                "label": "Office Professional 2024 Offline",
                "kind": "action",
                "action_id": "install_office_2024_prof_offline",
                "description": "Използва setup.exe и ConfigurationProPlus2024EnBgx64.xml от G:\\Installers\\Office 2024 Prof.",
            },
            {
                "label": "Office Standard 2024 Offline",
                "kind": "action",
                "action_id": "install_office_2024_standard_offline",
                "description": "Опитва инсталация от G:\\Installers\\Office 2024 Standart, ако файловете са налични.",
            },
            {
                "label": "Office Standard 2021 Offline",
                "kind": "action",
                "action_id": "install_office_2021_standard_offline",
                "description": "Опитва инсталация от G:\\Installers\\Office 2021 Standart, ако файловете са налични.",
            },
            {
                "label": "Office Online God Mode",
                "kind": "menu",
                "target": "office_center",
                "description": "Отваря online deployment и winget Office менюто.",
            },
            {"label": "Back to Install Software", "kind": "menu", "target": "install_software"},
        ],
    },
    "secret_install": {
        "title": "Secret Install Interface",
        "subtitle": "Grouped deployment presets for runtimes, tools and engineering stacks.",
        "items": [
            {"label": "System Runtimes", "kind": "action", "description": "Java, .NET, DirectX."},
            {"label": "Browsers & Comms", "kind": "action", "description": "Chrome, Discord, and communication tools."},
            {"label": "Development", "kind": "action", "description": "VS 2022, VS Code, Git, Docker."},
            {"label": "Languages & DB", "kind": "action", "description": "Python, Node, Java 21, SQL."},
            {"label": "Cybersecurity & Net", "kind": "action", "description": "Wireshark, Nmap, PuTTY."},
            {"label": "Virtualization", "kind": "action", "description": "VirtualBox, VMware."},
            {"label": "Multimedia & Design", "kind": "action", "description": "OBS, VLC, GIMP."},
            {"label": "Gaming & Tools", "kind": "action", "description": "Steam, Epic, DirectX."},
            {"label": "Utilities & Office", "kind": "action", "description": "7-Zip, LibreOffice, AnyDesk."},
            {"label": "Advanced Tools", "kind": "action", "description": "Sysinternals, Kali, scanners."},
            {"label": "Update All Apps", "kind": "action"},
            {"label": "Back to Install Software", "kind": "menu", "target": "install_software"},
        ],
    },
    "office_center": {
        "title": "Office Deployment Center",
        "subtitle": "Modern, legacy and maintenance tools for Microsoft Office ecosystems.",
        "items": [
            {"label": "Office Professional Plus 2024", "kind": "action", "action_id": "online_office_2024_proplus"},
            {"label": "Office Home & Business 2024", "kind": "action", "action_id": "online_office_2024_home_business"},
            {"label": "Office Professional Plus 2021", "kind": "action", "action_id": "online_office_2021_proplus"},
            {"label": "Office Home & Student 2021", "kind": "action", "action_id": "online_office_2021_home_student"},
            {"label": "Microsoft 365", "kind": "action", "action_id": "online_microsoft_365"},
            {"label": "Office Professional Plus 2019", "kind": "action", "action_id": "online_office_2019_proplus"},
            {"label": "Office Professional Plus 2016", "kind": "action", "action_id": "online_office_2016_proplus"},
            {"label": "Office Professional Plus 2013", "kind": "action", "action_id": "online_office_2013_proplus"},
            {"label": "Visio Professional 2024", "kind": "action", "action_id": "online_visio_2024_pro"},
            {"label": "Project Professional 2024", "kind": "action", "action_id": "online_project_2024_pro"},
            {"label": "Visio Professional 2021", "kind": "action", "action_id": "online_visio_2021_pro"},
            {"label": "Project Professional 2021", "kind": "action", "action_id": "online_project_2021_pro"},
            {
                "label": "Check Activation Status",
                "kind": "action",
                "action_id": "office_check_activation_status",
                "description": "Searches for OSPP.VBS and shows the Office activation status summary.",
            },
            {
                "label": "Quick Repair Office",
                "kind": "action",
                "action_id": "office_quick_repair",
                "description": "Launches Office Click-to-Run full repair when the repair tool is installed.",
            },
            {
                "label": "Force Uninstall All Office Versions",
                "kind": "action",
                "action_id": "office_force_uninstall_all",
                "description": "Runs the same winget cleanup flow from the batch script for all Office suites.",
            },
            {"label": "Back to Install Software", "kind": "menu", "target": "install_software"},
        ],
    },
    "language": {
        "title": "Windows 11 Language Manager",
        "subtitle": "Keyboard layouts and Bulgarian language pack options.",
        "items": [
            {
                "label": "Refresh Language Status",
                "kind": "action",
                "action_id": "language_refresh",
                "description": "Checks whether Bulgarian layouts and the language pack are already available.",
            },
            {
                "label": "Bulgarian BDS (Typewriter)",
                "kind": "action",
                "action_id": "toggle_bulgarian_bds",
                "description": "Adds the BDS keyboard if it is missing, or removes it if it is already present.",
            },
            {
                "label": "Bulgarian Phonetic (Standard)",
                "kind": "action",
                "action_id": "toggle_bulgarian_phonetic",
                "description": "Adds the standard phonetic layout or removes it if it is already installed.",
            },
            {
                "label": "Bulgarian Traditional Phonetic",
                "kind": "action",
                "action_id": "toggle_bulgarian_traditional",
                "description": "Adds the traditional phonetic layout or removes it if it is already installed.",
            },
            {
                "label": "Bulgarian Language Pack",
                "kind": "action",
                "action_id": "toggle_bulgarian_language_pack",
                "description": "Installs the Bulgarian language pack when missing, or removes it when already installed.",
            },
            {
                "label": "Remove Bulgarian Language Entry",
                "kind": "action",
                "action_id": "remove_bulgarian_language",
                "description": "Removes the bg-BG language entry from the current user language list.",
            },
            {"label": "Exit to Main Menu", "kind": "menu", "target": "main"},
        ],
    },
    "driver_backup": {
        "title": "Driver Backup God Mode",
        "subtitle": "Backup, recovery media and hardware reporting tools.",
        "items": [
            {
                "label": "Backup Drivers (Clean)",
                "kind": "action",
                "action_id": "driver_backup_clean",
                "description": "Recommended third-party driver export using pnputil, plus log, driver list and ZIP.",
            },
            {
                "label": "Backup Drivers (Full)",
                "kind": "action",
                "action_id": "driver_backup_full",
                "description": "Full DISM driver export with log, driver list and ZIP archive.",
            },
            {
                "label": "Create Recovery USB + RESTORE.bat",
                "kind": "action",
                "action_id": "driver_recovery_usb",
                "description": "Copies the last backup to a removable USB drive and creates RESTORE_DRIVERS.bat.",
            },
            {
                "label": "Generate PC Report",
                "kind": "action",
                "action_id": "driver_pc_report",
                "description": "Creates a Speccy-like hardware report with CPU, RAM, GPU, BIOS, disks and network info.",
            },
            {
                "label": "Driver Backup Tool v0.1",
                "kind": "action",
                "action_id": "driver_backup_advanced",
                "description": "Advanced mode with destination choice, backup type, ZIP options and restore workflow.",
            },
            {
                "label": "Restore Drivers From Last Backup",
                "kind": "action",
                "action_id": "driver_restore_last",
                "description": "Useful extra: reinstalls drivers directly from the last saved backup folder.",
            },
            {"label": "Return to Main Menu", "kind": "menu", "target": "main"},
        ],
    },
    "nexus_admin": {
        "title": "System Commander: Nexus Admin",
        "subtitle": "Local account management based on the batch menu, plus a few useful admin extras.",
        "items": [
            {
                "label": "List All Users",
                "kind": "action",
                "action_id": "nexus_list_users",
                "description": "Shows all local users on this PC with enabled state and last logon information.",
            },
            {
                "label": "Change Password",
                "kind": "action",
                "action_id": "nexus_change_password",
                "description": "Changes the password of an existing local user.",
            },
            {
                "label": "Create New User",
                "kind": "action",
                "action_id": "nexus_create_user",
                "description": "Creates a local account, with optional password and optional Administrator rights.",
            },
            {
                "label": "Delete User",
                "kind": "action",
                "action_id": "nexus_delete_user",
                "description": "Permanently removes a local account after confirmation.",
            },
            {
                "label": "User Details",
                "kind": "action",
                "action_id": "nexus_user_details",
                "description": "Useful extra: shows full `net user` details for one selected account.",
            },
            {
                "label": "Toggle Administrator Rights",
                "kind": "action",
                "action_id": "nexus_toggle_admin",
                "description": "Useful extra: adds or removes a user from the local Administrators group.",
            },
            {"label": "Return to Main Menu", "kind": "menu", "target": "main"},
        ],
    },
}

MENU_LABELS_TO_REMOVE = {
    "Return to Main Menu",
    "Back to Install Software",
    "Exit to Main Menu",
}

UI_TRANSLATIONS = {
    "Main Menu": "Главно меню",
    "Central control hub for deployment, activation, language, recovery and admin tools.": "Централен контролен панел за активация, инсталации, езикови настройки, архивиране и администраторски инструменти.",
    "Activation Menu": "Меню за активация",
    "Add Desktop Icons": "Добави икони на работния плот",
    "Create standard support shortcuts.": "Създава стандартни преки пътища за поддръжка и бърз достъп.",
    "Reset OneDrive": "Нулиране на OneDrive",
    "Install Software": "Инсталиране на софтуер",
    "Language Menu": "Езиково меню",
    "Driver Backup + PC Report": "Архив на драйвери и отчет за компютъра",
    "System Commander: Nexus Admin": "Системен командир: Nexus Admin",
    "Local user management and administrator account tools.": "Управление на локални потребители и администраторски акаунти.",
    "Reset Console": "Нулирай конзолата",
    "Refresh the current interface state.": "Освежава текущото състояние на интерфейса.",
    "Exit": "Изход",
    "Close WinSys Guardian Advanced.": "Затваря WinSys Guardian Advanced.",
    "Windows and Office activation shortcuts.": "Бърз достъп до активация на Windows и Office.",
    "Activate Windows 10": "Активирай Windows 10",
    "Activate Windows 11": "Активирай Windows 11",
    "Office Activation Center": "Център за активация на Office",
    "Windows 10 Key Manager": "Управление на ключ за Windows 10",
    "Save and manage the Windows 10 product key used by your admin team.": "Запис и управление на ключа за Windows 10, използван от администратора.",
    "Run Windows 10 Activation": "Стартирай активация на Windows 10",
    "Store the currently approved Windows 10 key for later use.": "Записва текущия одобрен ключ за Windows 10 за по-късно използване.",
    "Windows 11 Key Manager": "Управление на ключ за Windows 11",
    "Save and manage the Windows 11 product key used by your admin team.": "Запис и управление на ключа за Windows 11, използван от администратора.",
    "Run Windows 11 Activation": "Стартирай активация на Windows 11",
    "Run the Windows activation commands using the saved product key.": "Изпълнява командите за активация на Windows със записания продуктов ключ.",
    "Save or Replace Product Key": "Запази или смени продуктов ключ",
    "Store the currently approved Windows 11 key for later use.": "Записва текущия одобрен ключ за Windows 11 за по-късно използване.",
    "Show Saved Product Key": "Покажи записания продуктов ключ",
    "Display the key currently stored in this application.": "Показва ключа, който в момента е записан в приложението.",
    "Clear Saved Product Key": "Изтрий записания продуктов ключ",
    "Remove the saved key if your organization replaces it.": "Премахва записания ключ, ако бъде заменен.",
    "Save and manage the Office product key, then choose the target Office version.": "Запис и управление на ключ за Office, след което избор на версия за активация.",
    "Save or Replace Office Key": "Запази или смени Office ключ",
    "Store the Office product key used by your admin workflow.": "Записва продуктовия ключ за Office, който използваш в администрацията.",
    "Show Saved Office Key": "Покажи записания Office ключ",
    "Display the Office key currently stored in this application.": "Показва записания в приложението ключ за Office.",
    "Clear Saved Office Key": "Изтрий записания Office ключ",
    "Remove the saved Office key if your organization replaces it.": "Премахва записания ключ за Office, ако бъде заменен.",
    "Run activation workflow for Office 2016 using the saved Office key.": "Стартира активация на Office 2016 със записания ключ.",
    "Run activation workflow for Office 2019 using the saved Office key.": "Стартира активация на Office 2019 със записания ключ.",
    "Run activation workflow for Office 2021 using the saved Office key.": "Стартира активация на Office 2021 със записания ключ.",
    "Choose a reset workflow for the OneDrive client.": "Избери метод за нулиране на OneDrive клиента.",
    "Reset OneDrive (Method 1)": "Нулиране на OneDrive (Метод 1)",
    "Reset OneDrive (Method 2)": "Нулиране на OneDrive (Метод 2)",
    "Reset OneDrive (Method 3)": "Нулиране на OneDrive (Метод 3)",
    "Install Software Menu": "Меню за инсталиране",
    "Office installers, app deployment and advanced install hubs.": "Инсталатори за Office, приложения и разширени менюта за инсталация.",
    "Office Install Center": "Център за инсталиране на Office",
    "Office 2016 Offline": "Office 2016 локално",
    "Office 2019 Offline": "Office 2019 локално",
    "Office 2021 Offline": "Office 2021 локално",
    "Office Professional 2021 Offline": "Office Professional 2021 локално",
    "Office Professional 2024 Offline": "Office Professional 2024 локално",
    "Office Standard 2024 Offline": "Office Standard 2024 локално",
    "Office Standard 2021 Offline": "Office Standard 2021 локално",
    "Office Online God Mode": "Office онлайн God Mode",
    "Install Ninite": "Инсталирай Ninite",
    "Install Adobe Reader": "Инсталирай Adobe Reader",
    "Secret Install Interface": "Скрито меню за инсталации",
    "Grouped deployment presets for runtimes, tools and engineering stacks.": "Групирани категории за инсталиране на среди, инструменти и специализиран софтуер.",
    "System Runtimes": "Системни среди",
    "Browsers & Comms": "Браузъри и комуникация",
    "Development": "Разработка",
    "Languages & DB": "Езици и бази данни",
    "Cybersecurity & Net": "Киберсигурност и мрежи",
    "Virtualization": "Виртуализация",
    "Multimedia & Design": "Мултимедия и дизайн",
    "Gaming & Tools": "Игри и инструменти",
    "Utilities & Office": "Полезни програми и офис",
    "Advanced Tools": "Разширени инструменти",
    "Update All Apps": "Обнови всички приложения",
    "Office Deployment Center": "Център за внедряване на Office",
    "Modern, legacy and maintenance tools for Microsoft Office ecosystems.": "Модерни, стари и сервизни инструменти за Microsoft Office.",
    "Check Activation Status": "Провери статуса на активацията",
    "Quick Repair Office": "Бърз ремонт на Office",
    "Force Uninstall All Office Versions": "Принудително премахни всички версии на Office",
    "Searches for OSPP.VBS and shows the Office activation status summary.": "Търси OSPP.VBS и показва обобщен статус на активацията на Office.",
    "Launches Office Click-to-Run full repair when the repair tool is installed.": "Стартира пълния ремонт на Office, ако инструментът за ремонт е наличен.",
    "Runs the same winget cleanup flow from the batch script for all Office suites.": "Изпълнява същото winget почистване от batch файла за всички Office пакети.",
    "Windows 11 Language Manager": "Езиков мениджър за Windows 11",
    "Keyboard layouts and Bulgarian language pack options.": "Управление на клавиатурни подредби и български езиков пакет.",
    "Refresh Language Status": "Обнови езиковия статус",
    "Checks whether Bulgarian layouts and the language pack are already available.": "Проверява дали българските подредби и езиковият пакет вече са налични.",
    "Bulgarian BDS (Typewriter)": "Български БДС",
    "Adds the BDS keyboard if it is missing, or removes it if it is already present.": "Добавя БДС подредбата, ако липсва, или я премахва, ако вече е налична.",
    "Bulgarian Phonetic (Standard)": "Български фонетичен",
    "Adds the standard phonetic layout or removes it if it is already installed.": "Добавя стандартната фонетична подредба или я премахва, ако вече е налична.",
    "Bulgarian Traditional Phonetic": "Български традиционен фонетичен",
    "Adds the traditional phonetic layout or removes it if it is already installed.": "Добавя традиционната фонетична подредба или я премахва, ако вече е налична.",
    "Bulgarian Language Pack": "Български езиков пакет",
    "Installs the Bulgarian language pack when missing, or removes it when already installed.": "Инсталира българския езиков пакет, ако липсва, или го премахва, ако е наличен.",
    "Remove Bulgarian Language Entry": "Премахни българския език от списъка",
    "Removes the bg-BG language entry from the current user language list.": "Премахва записа `bg-BG` от текущия езиков списък на потребителя.",
    "Driver Backup God Mode": "Driver Backup God Mode",
    "Backup, recovery media and hardware reporting tools.": "Архивиране на драйвери, създаване на recovery носител и хардуерен отчет.",
    "Backup Drivers (Clean)": "Архив на драйвери (чист)",
    "Recommended third-party driver export using pnputil, plus log, driver list and ZIP.": "Препоръчителен архив само на външните драйвери чрез pnputil, с лог, списък и ZIP.",
    "Backup Drivers (Full)": "Архив на драйвери (пълен)",
    "Full DISM driver export with log, driver list and ZIP archive.": "Пълен експорт на драйвери чрез DISM, с лог, списък и ZIP архив.",
    "Create Recovery USB + RESTORE.bat": "Създай Recovery USB + RESTORE.bat",
    "Copies the last backup to a removable USB drive and creates RESTORE_DRIVERS.bat.": "Копира последния архив на USB устройство и създава RESTORE_DRIVERS.bat.",
    "Generate PC Report": "Генерирай отчет за компютъра",
    "Creates a Speccy-like hardware report with CPU, RAM, GPU, BIOS, disks and network info.": "Създава подробен хардуерен отчет с процесор, RAM, видео, BIOS, дискове и мрежа.",
    "Driver Backup Tool v0.1": "Driver Backup Tool v0.1",
    "Advanced mode with destination choice, backup type, ZIP options and restore workflow.": "Разширен режим с избор на дестинация, тип архив, ZIP настройки и възстановяване.",
    "Restore Drivers From Last Backup": "Възстанови драйверите от последния архив",
    "Useful extra: reinstalls drivers directly from the last saved backup folder.": "Полезна екстра: преинсталира драйверите директно от последната записана архивна папка.",
    "Create Recovery USB + RESTORE.bat": "Създай Recovery USB + RESTORE.bat",
    "Local account management based on the batch menu, plus a few useful admin extras.": "Управление на локални акаунти по batch менюто, плюс няколко полезни админ екстри.",
    "List All Users": "Покажи всички потребители",
    "Shows all local users on this PC with enabled state and last logon information.": "Показва всички локални потребители с активност и последно влизане.",
    "Change Password": "Смени парола",
    "Changes the password of an existing local user.": "Променя паролата на съществуващ локален потребител.",
    "Create New User": "Създай нов потребител",
    "Creates a local account, with optional password and optional Administrator rights.": "Създава локален акаунт с опционална парола и опционални администраторски права.",
    "Delete User": "Изтрий потребител",
    "Permanently removes a local account after confirmation.": "Изтрива локален акаунт след потвърждение.",
    "User Details": "Детайли за потребител",
    "Useful extra: shows full `net user` details for one selected account.": "Полезна екстра: показва пълните `net user` детайли за избран акаунт.",
    "Toggle Administrator Rights": "Промени администраторските права",
    "Useful extra: adds or removes a user from the local Administrators group.": "Полезна екстра: добавя или премахва потребител от локалната група Administrators.",
}


# Помощна функция за localize menu tree.
def _localize_menu_tree(data: object) -> object:
    if isinstance(data, dict):
        localized: dict[str, object] = {}
        for key, value in data.items():
            if key == "items" and isinstance(value, list):
                filtered_items = []
                for item in value:
                    if isinstance(item, dict) and item.get("label") in MENU_LABELS_TO_REMOVE:
                        continue
                    filtered_items.append(_localize_menu_tree(item))
                localized[key] = filtered_items
            elif isinstance(value, str):
                localized[key] = UI_TRANSLATIONS.get(value, value)
            else:
                localized[key] = _localize_menu_tree(value)
        return localized
    if isinstance(data, list):
        return [_localize_menu_tree(item) for item in data]
    return data


MENU_TREE = _localize_menu_tree(MENU_TREE)


FILE_ATTRIBUTE_HIDDEN = 0x02
FILE_ATTRIBUTE_NORMAL = 0x80


# Зарежда settings от файл или конфигурация.
def load_settings() -> dict[str, str]:
    # Зарежда обикновените настройки на приложението.
    if not SETTINGS_FILE.exists():
        return {}
    try:
        data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}
    return data if isinstance(data, dict) else {}


# Записва settings за следващо използване.
def save_settings(settings: dict[str, str]) -> None:
    # Записва настройките в settings.json.
    SETTINGS_FILE.write_text(json.dumps(settings, indent=2), encoding="utf-8")


# Зарежда version info от файл или конфигурация.
def load_version_info() -> dict[str, object]:
    # Зарежда локалната версия и адресите за online update.
    defaults = {
        "version": "0.1.1",
        "version_info_url": "",
        "download_url": "",
        "package_url": "",
        "notes": "",
        "changelog": [],
    }
    if not VERSION_FILE.exists():
        return defaults
    try:
        data = json.loads(VERSION_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return defaults
    if not isinstance(data, dict):
        return defaults
    merged = defaults.copy()
    for key, value in data.items():
        if value is None:
            continue
        if key == "changelog" and isinstance(value, list):
            merged[key] = [str(item) for item in value if str(item).strip()]
        else:
            merged[key] = str(value)
    return merged


# Помощна функция за format bytes per second.
def format_bytes_per_second(value: float) -> str:
    # Форматира скоростта в удобен за четене вид.
    units = ("B/s", "KB/s", "MB/s", "GB/s")
    current = float(max(0.0, value))
    for unit in units:
        if current < 1024 or unit == units[-1]:
            return f"{current:.1f} {unit}"
        current /= 1024
    return f"{current:.1f} GB/s"


# Помощна функция за format file size.
def format_file_size(value: int) -> str:
    # Форматира размер на файл в B, KB, MB или GB.
    units = ("B", "KB", "MB", "GB")
    current = float(max(0, value))
    for unit in units:
        if current < 1024 or unit == units[-1]:
            return f"{current:.1f} {unit}"
        current /= 1024
    return f"{current:.1f} GB"


# Помощна функция за format duration.
def format_duration(seconds: int) -> str:
    # Форматира секунди като оставащо време.
    seconds = max(0, int(seconds))
    hours, remainder = divmod(seconds, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours:
        return f"{hours} ч {minutes:02d} мин"
    if minutes:
        return f"{minutes} мин {secs:02d} сек"
    return f"{secs} сек"


# Помощна функция за portable secret key.
def _portable_secret_key() -> bytes:
    # Прави локален ключ за леко скриване на чувствителните данни.
    secret_seed = f"{APP_TITLE}|WGA-Portable-Store|{PROJECT_ROOT.name}"
    return hashlib.sha256(secret_seed.encode("utf-8")).digest()


# Помощна функция за encrypt for current user.
def encrypt_for_current_user(text: str) -> str:
    # Кодира текста преди да се запише в защитения файл.
    source = text.encode("utf-8")
    key = _portable_secret_key()
    encrypted = bytes(byte ^ key[index % len(key)] for index, byte in enumerate(source))
    return base64.b64encode(encrypted).decode("ascii")


# Помощна функция за decrypt for current user.
def decrypt_for_current_user(encoded_text: str) -> str:
    # Декодира текста, записан в защитения файл.
    encrypted = base64.b64decode(encoded_text.encode("ascii"))
    key = _portable_secret_key()
    decrypted = bytes(byte ^ key[index % len(key)] for index, byte in enumerate(encrypted))
    return decrypted.decode("utf-8")


# Помощна функция за hash secret.
def hash_secret(value: str) -> str:
    # Прави hash на парола или ключ, без да пазим оригинала.
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


# Помощна функция за ensure hidden file.
def ensure_hidden_file(path: Path) -> None:
    # Скрива файла в Windows Explorer.
    ctypes.windll.kernel32.SetFileAttributesW(str(path), FILE_ATTRIBUTE_HIDDEN)


# Помощна функция за ensure normal file.
def ensure_normal_file(path: Path) -> None:
    # Връща файла в нормален вид, за да може да се редактира или замени.
    if path.exists():
        ctypes.windll.kernel32.SetFileAttributesW(str(path), FILE_ATTRIBUTE_NORMAL)


# Връща drive label в удобен за останалия код вид.
def get_drive_label(path: Path) -> str:
    # Връща името на устройството, например името на флашката.
    drive_root = path.anchor or str(path.drive)
    if not drive_root:
        return "Unknown"

    volume_name = ctypes.create_unicode_buffer(261)
    filesystem_name = ctypes.create_unicode_buffer(261)
    result = ctypes.windll.kernel32.GetVolumeInformationW(
        ctypes.c_wchar_p(drive_root),
        volume_name,
        ctypes.sizeof(volume_name),
        None,
        None,
        None,
        filesystem_name,
        ctypes.sizeof(filesystem_name),
    )
    if not result:
        return "Unnamed Drive"
    return volume_name.value or "Unnamed Drive"


# Връща launch location info в удобен за останалия код вид.
def get_launch_location_info() -> dict[str, str]:
    # Събира информация откъде е стартирано приложението.
    storage_info = get_runtime_storage_info(PROJECT_ROOT)
    return {
        "program_path": str(PROJECT_ROOT),
        "drive": storage_info.drive or "Unknown",
        "device_name": get_drive_label(PROJECT_ROOT),
        "drive_type": storage_info.drive_type,
        "drive_type_label": storage_info.drive_type_label,
        "installers_root": str(storage_info.installers_root),
        "installers_available": "Yes" if storage_info.installers_available else "No",
    }


# Зарежда secure store от файл или конфигурация.
def load_secure_store() -> dict[str, str]:
    # Зарежда скрития файл с ключове и служебни пароли.
    if not SECURE_STORE_FILE.exists():
        store = {"admin_menu_password_hash": hash_secret(DEFAULT_WINDOWS11_MENU_PASSWORD)}
        save_secure_store(store)
        return store

    try:
        encrypted_payload = json.loads(SECURE_STORE_FILE.read_text(encoding="utf-8"))
        encrypted_data = encrypted_payload.get("data", "")
        if not encrypted_data:
            raise ValueError("Missing encrypted data.")
        decrypted = decrypt_for_current_user(encrypted_data)
        data = json.loads(decrypted)
    except (OSError, ValueError, json.JSONDecodeError):
        data = {"admin_menu_password_hash": hash_secret(DEFAULT_WINDOWS11_MENU_PASSWORD)}
        try:
            ensure_normal_file(SECURE_STORE_FILE)
            if SECURE_STORE_FILE.exists():
                backup_file = SECURE_STORE_FILE.with_suffix(".json.bak")
                if backup_file.exists():
                    ensure_normal_file(backup_file)
                    backup_file.unlink()
                SECURE_STORE_FILE.replace(backup_file)
        except OSError:
            pass
        save_secure_store(data)

    return data if isinstance(data, dict) else {"admin_menu_password_hash": hash_secret(DEFAULT_WINDOWS11_MENU_PASSWORD)}


# Записва secure store за следващо използване.
def save_secure_store(store: dict[str, str]) -> None:
    # Записва защитените данни обратно в скрития файл.
    serialized = json.dumps(store, indent=2)
    encrypted = encrypt_for_current_user(serialized)
    payload = {"data": encrypted}
    ensure_normal_file(SECURE_STORE_FILE)
    SECURE_STORE_FILE.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    ensure_hidden_file(SECURE_STORE_FILE)


# Помощна функция за is running as admin.
def is_running_as_admin() -> bool:
    # Проверява дали приложението има администраторски права.
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except OSError:
        return False


# Помощна функция за apply app icon.
def apply_app_icon(root: tk.Tk | tk.Toplevel) -> None:
    # Слага иконата на прозорците на приложението.
    if not APP_ICON_FILE.exists():
        return
    try:
        root.iconbitmap(str(APP_ICON_FILE))
    except tk.TclError:
        pass


# Помощна функция за relaunch as admin.
def relaunch_as_admin() -> bool:
    # Стартира приложението отново с admin права и запазва подадените параметри.
    current_args = sys.argv[1:]
    if getattr(sys, "frozen", False):
        target_path = str(Path(sys.executable).resolve())
        run_args = subprocess.list2cmdline(current_args)
    else:
        script_path = str(Path(__file__).resolve())
        target_path = sys.executable
        run_args = subprocess.list2cmdline([script_path, *current_args])
    result = ctypes.windll.shell32.ShellExecuteW(
        None,
        "runas",
        target_path,
        run_args,
        None,
        1,
    )
    return result > 32


# Връща startup menu from args в удобен за останалия код вид.
def get_startup_menu_from_args() -> str | None:
    # Чете подаденото меню от shortcut аргумент като: --menu windows11_activation
    args = sys.argv[1:]
    for index, value in enumerate(args):
        if value == "--menu" and index + 1 < len(args):
            menu_key = args[index + 1].strip()
            if menu_key in MENU_TREE:
                return menu_key
    return None


# Помощна функция за enable windows desktop icons.
def enable_windows_desktop_icons(progress_callback=None) -> list[str]:
    # Показва системните икони на работния плот през Windows registry.
    enabled_labels: list[str] = []
    total_steps = len(DESKTOP_ICON_TARGETS) + 2
    current_step = 0

    for label, clsid in DESKTOP_ICON_TARGETS:
        current_step += 1
        progress_value = min(90, int(current_step / total_steps * 100))
        if progress_callback is not None:
            progress_callback(
                progress_value,
                f"Активиране на {label}...",
                f"Включва се системната икона {label} на работния плот.",
            )
        for registry_path in DESKTOP_ICON_PATHS:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, registry_path)
            winreg.SetValueEx(key, clsid, 0, winreg.REG_DWORD, 0)
            winreg.CloseKey(key)
        enabled_labels.append(label)

    current_step += 1
    if progress_callback is not None:
        progress_callback(94, "Опресняване на работния плот...", "Explorer се опреснява, за да се покажат иконите.")
    refresh_windows_desktop()

    current_step += 1
    if progress_callback is not None:
        progress_callback(100, "Готово.", "Иконите на работния плот са активирани успешно.")
    return enabled_labels


# Помощна функция за refresh windows desktop.
def refresh_windows_desktop() -> None:
    # Опреснява работния плот след промяна на системните икони.
    try:
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)
    except Exception:
        pass

    refresh_commands = [
        ["ie4uinit.exe", "-show"],
        ["rundll32.exe", "user32.dll,UpdatePerUserSystemParameters"],
    ]
    for command in refresh_commands:
        try:
            subprocess.run(command, check=False, capture_output=True, text=True)
        except OSError:
            continue


# Помощна функция за switch keyboard layout to english.
def switch_keyboard_layout_to_english() -> int | None:
    # Превключва клавиатурата към английска подредба за по-сигурно въвеждане на ключове.
    user32 = ctypes.windll.user32
    current_layout = user32.GetKeyboardLayout(0)
    english_layout = user32.LoadKeyboardLayoutW("00000409", 1)
    if english_layout:
        user32.ActivateKeyboardLayout(english_layout, 0)
    return current_layout if current_layout else None


# Възстановява keyboard layout от подготвен backup.
def restore_keyboard_layout(layout_handle: int | None) -> None:
    # Връща старата клавиатурна подредба след като приключим с въвеждането.
    if layout_handle:
        ctypes.windll.user32.ActivateKeyboardLayout(layout_handle, 0)


# Помощна функция за normalize product key input.
def normalize_product_key_input(raw_value: str) -> str:
    # Нормализира продуктов ключ и поправя често срещано въвеждане на кирилица.
    bg_to_latin = str.maketrans(
        {
            "А": "A",
            "Р’": "B",
            "С": "C",
            "Р•": "E",
            "Н": "H",
            "К": "K",
            "М": "M",
            "О": "O",
            "Р ": "P",
            "Т": "T",
            "Х": "X",
            "У": "Y",
            "Р°": "A",
            "в": "B",
            "с": "C",
            "Рµ": "E",
            "н": "H",
            "к": "K",
            "м": "M",
            "о": "O",
            "р": "P",
            "С‚": "T",
            "С…": "X",
            "у": "Y",
            "Р¬": "B",
            "ь": "B",
            "Р†": "I",
            "С–": "I",
        }
    )
    normalized = raw_value.strip().translate(bg_to_latin).upper()
    allowed_chars = set("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-")
    normalized = "".join(char for char in normalized if char in allowed_chars)
    return normalized


# Помощна функция за ask product key.
def ask_product_key(parent: tk.Misc, title: str, prompt: str, initialvalue: str = "") -> str | None:
    # Показва поле за ключ, като първо превключва към английска клавиатура.
    previous_layout = switch_keyboard_layout_to_english()
    try:
        return simpledialog.askstring(
            title,
            prompt,
            parent=parent,
            initialvalue=initialvalue,
        )
    finally:
        restore_keyboard_layout(previous_layout)


# Начален launcher, който избира кой WGA модул да бъде стартиран.
class ProductLauncher:
    MODULES: tuple[dict[str, str], ...] = (
        {
            "id": "wga",
            "title": "WGA",
            "subtitle": "Системна поддръжка, инсталации и администрация",
            "icon": "assets/wga-icon.png",
            "accent": "#37e39a",
        },
        {
            "id": "optimization",
            "title": "ОПТИМИЗАЦИЯ",
            "subtitle": "Почистване, настройване и ускоряване на Windows",
            "icon": "assets/dashboard-icons/dashboard-bolt.png",
            "accent": "#d0a94a",
        },
        {
            "id": "network",
            "title": "NETWORK MONITORING",
            "subtitle": "Наблюдение на връзката и мрежовите устройства",
            "icon": "assets/dashboard-icons/dashboard-monitor.png",
            "accent": "#2f8fff",
        },
    )

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("WinSys Guardian Suite")
        apply_app_icon(self.root)
        apply_tk_dpi_scaling(self.root)
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width, height = responsive_window_size(screen_width, screen_height, 1040, 690)
        center_geometry(self.root, width, height)
        self.root.configure(bg="#071311")
        self.root.minsize(min(width, 900), min(height, 540))
        self.root.resizable(False, False)
        self.root.overrideredirect(True)
        self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)
        self.module_icons: list[tk.PhotoImage] = []
        self.drag_offset_x = 0
        self.drag_offset_y = 0
        self.version_info = load_version_info()
        self.update_result: UpdateResult | None = None
        self.update_check_active = True
        self.update_status_var = tk.StringVar(
            value=f"Проверка за актуализация на v{self.version_info['version']}..."
        )
        self._build_interface()
        self._check_updates_async()

    def _build_interface(self) -> None:
        shell = tk.Frame(
            self.root,
            bg="#071311",
            highlightbackground="#1f4e46",
            highlightthickness=1,
        )
        shell.pack(fill="both", expand=True, padx=22, pady=22)

        header = tk.Frame(shell, bg="#0a1b18", height=118)
        header.pack(fill="x")
        header.pack_propagate(False)
        header.bind("<ButtonPress-1>", self._start_window_drag)
        header.bind("<B1-Motion>", self._drag_window)

        title_label = tk.Label(
            header,
            text="WinSys Guardian",
            bg="#0a1b18",
            fg="#ecfff7",
            font=("Segoe UI Semibold", 28),
        )
        title_label.pack(pady=(22, 0))
        title_label.bind("<ButtonPress-1>", self._start_window_drag)
        title_label.bind("<B1-Motion>", self._drag_window)
        subtitle_label = tk.Label(
            header,
            text="ИЗБЕРЕТЕ МОДУЛ",
            bg="#0a1b18",
            fg="#37e39a",
            font=("Segoe UI Semibold", 10),
        )
        subtitle_label.pack(pady=(3, 0))
        subtitle_label.bind("<ButtonPress-1>", self._start_window_drag)
        subtitle_label.bind("<B1-Motion>", self._drag_window)

        tk.Button(
            header,
            text="×",
            command=self.root.destroy,
            bg="#0a1b18",
            activebackground="#c94d58",
            fg="#7ca394",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Segoe UI Semibold", 17),
            width=3,
        ).place(relx=1.0, x=-10, y=8, anchor="ne")

        content = tk.Frame(shell, bg="#071311")
        content.pack(fill="both", expand=True, padx=26, pady=(32, 22))

        tk.Label(
            content,
            text="Център за системно управление",
            bg="#071311",
            fg="#ecfff7",
            font=("Segoe UI Semibold", 18),
        ).pack()
        tk.Label(
            content,
            text="Изберете инструмент, за да продължите",
            bg="#071311",
            fg="#7ca394",
            font=("Segoe UI", 10),
        ).pack(pady=(4, 26))

        cards = tk.Frame(content, bg="#071311")
        cards.pack(fill="both", expand=True)
        for column in range(len(self.MODULES)):
            cards.grid_columnconfigure(column, weight=1, uniform="launcher-card")
        cards.grid_rowconfigure(0, weight=1)

        for column, module in enumerate(self.MODULES):
            card = tk.Frame(
                cards,
                bg="#0d1c1a",
                highlightbackground=str(module["accent"]),
                highlightthickness=1,
                cursor="hand2",
            )
            card.grid(row=0, column=column, sticky="nsew", padx=9)

            icon = self._load_module_icon(str(module["icon"]), 78)
            if icon is not None:
                self.module_icons.append(icon)
                tk.Label(card, image=icon, bg="#0d1c1a").pack(pady=(162, 13))
            else:
                tk.Label(
                    card,
                    text="◆",
                    bg="#0d1c1a",
                    fg=str(module["accent"]),
                    font=("Segoe UI", 42),
                ).pack(pady=(162, 13))

            tk.Label(
                card,
                text=str(module["title"]),
                bg="#0d1c1a",
                fg="#ecfff7",
                font=("Segoe UI Semibold", 14),
            ).pack()

            action = lambda module_id=str(module["id"]): self._open_module(module_id)
            button = tk.Button(
                card,
                text="СТАРТИРАЙ",
                command=action,
                bg=str(module["accent"]),
                activebackground="#ecfff7",
                fg="#071311",
                activeforeground="#071311",
                relief="flat",
                bd=0,
                cursor="hand2",
                font=("Segoe UI Semibold", 10),
                padx=24,
                pady=10,
            )
            button.pack(pady=(20, 16))

            description = tk.Label(
                card,
                text=str(module["subtitle"]),
                bg="#0d1c1a",
                fg="#a6d5c5",
                justify="center",
                wraplength=220,
                font=("Segoe UI", 9),
            )
            description.pack(padx=18, pady=(0, 10))
            for widget in card.winfo_children():
                if widget is not button:
                    widget.bind("<Button-1>", lambda _event, callback=action: callback())
            card.bind("<Button-1>", lambda _event, callback=action: callback())

        update_bar = tk.Frame(
            shell,
            bg="#0d1c1a",
            highlightbackground="#1f4e46",
            highlightthickness=1,
        )
        update_bar.pack(fill="x", padx=26, pady=(0, 10))

        self.update_status_icon = tk.Label(
            update_bar,
            text="↻",
            bg="#0d1c1a",
            fg="#37e39a",
            font=("Segoe UI Semibold", 14),
            width=3,
        )
        self.update_status_icon.pack(side="left", padx=(8, 2), pady=8)

        self.update_status_label = tk.Label(
            update_bar,
            textvariable=self.update_status_var,
            bg="#0d1c1a",
            fg="#a6d5c5",
            anchor="w",
            font=("Segoe UI", 9),
        )
        self.update_status_label.pack(side="left", fill="x", expand=True, pady=8)

        tk.Button(
            update_bar,
            text="История на актуализациите",
            command=self._show_update_history,
            bg="#173c4d",
            activebackground="#1b5d73",
            fg="#f3fbff",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Segoe UI Semibold", 9),
            padx=14,
            pady=7,
        ).pack(side="right", padx=8, pady=6)

        tk.Label(
            shell,
            text=f"WinSys Guardian Advanced v{self.version_info['version']} • Administrative Toolkit",
            bg="#071311",
            fg="#557b70",
            font=("Segoe UI", 9),
        ).pack(pady=(0, 14))

    def _load_module_icon(self, relative_path: str, target_size: int) -> tk.PhotoImage | None:
        icon_path = runtime_file(relative_path)
        if not icon_path.exists():
            return None
        try:
            image = tk.PhotoImage(file=str(icon_path))
            largest_side = max(image.width(), image.height())
            factor = max(1, math.ceil(largest_side / target_size))
            return image.subsample(factor, factor)
        except tk.TclError:
            return None

    def _start_window_drag(self, event: tk.Event) -> None:
        self.drag_offset_x = event.x_root - self.root.winfo_x()
        self.drag_offset_y = event.y_root - self.root.winfo_y()

    def _drag_window(self, event: tk.Event) -> None:
        self.root.geometry(
            f"+{event.x_root - self.drag_offset_x}+{event.y_root - self.drag_offset_y}"
        )

    def _check_updates_async(self) -> None:
        threading.Thread(target=self._perform_update_check, daemon=True).start()

    def _perform_update_check(self) -> None:
        result = check_for_updates(
            str(self.version_info["version"]),
            str(self.version_info.get("version_info_url", "")),
        )
        try:
            self.root.after(0, lambda: self._apply_update_result(result))
        except tk.TclError:
            return

    def _apply_update_result(self, result: UpdateResult) -> None:
        self.update_result = result
        if not self.update_check_active:
            return
        try:
            if not self.update_status_label.winfo_exists():
                return
        except tk.TclError:
            return

        if result.status == "up_to_date":
            icon, color = "✓", "#37e39a"
            message = f"Приложението е актуално — версия v{self.version_info['version']}."
        elif result.status == "update_available":
            icon, color = "↑", "#d0a94a"
            message = f"Налична е версия v{result.latest_version}. {result.notes or ''}".strip()
        elif result.status in {"not_configured", "raw_unavailable"}:
            icon, color = "!", "#d0a94a"
            message = "Онлайн проверката за актуализации не е достъпна."
        else:
            icon, color = "×", "#c94d58"
            message = f"Проверката за актуализация не успя: {result.error or 'неизвестна грешка'}"

        self.update_status_icon.configure(text=icon, fg=color)
        self.update_status_label.configure(fg=color)
        self.update_status_var.set(message)

    def _update_history_lines(self) -> list[str]:
        if self.update_result and self.update_result.changelog:
            return list(self.update_result.changelog)
        raw_changelog = self.version_info.get("changelog", [])
        if isinstance(raw_changelog, list):
            return [str(item) for item in raw_changelog if str(item).strip()]
        return []

    def _show_update_history(self) -> None:
        history_window = tk.Toplevel(self.root)
        history_window.title("История на актуализациите")
        history_window.geometry("680x470")
        history_window.transient(self.root)
        history_window.configure(bg="#0d1711")
        apply_app_icon(history_window)

        wrapper = tk.Frame(history_window, bg="#0d1711", padx=20, pady=18)
        wrapper.pack(fill="both", expand=True)
        tk.Label(
            wrapper,
            text="История на актуализациите",
            font=("Segoe UI Semibold", 16),
            bg="#0d1711",
            fg="#effff2",
        ).pack(anchor="w")
        tk.Label(
            wrapper,
            text=self.update_status_var.get(),
            font=("Segoe UI", 10),
            bg="#0d1711",
            fg="#aee8b8",
            wraplength=620,
            justify="left",
        ).pack(anchor="w", pady=(5, 14))

        text_box = tk.Text(
            wrapper,
            bg="#07100a",
            fg="#e7ffe9",
            insertbackground="#e7ffe9",
            relief="flat",
            wrap="word",
            font=("Segoe UI", 10),
            padx=14,
            pady=14,
        )
        text_box.pack(fill="both", expand=True)
        lines = self._update_history_lines()
        text_box.insert(
            "end",
            "\n\n".join(lines) if lines else "Все още няма добавена история на актуализациите.",
        )
        text_box.configure(state="disabled")
        tk.Button(
            wrapper,
            text="Затвори",
            command=history_window.destroy,
            bg="#245634",
            activebackground="#2f7044",
            fg="#f3fff5",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Segoe UI Semibold", 10),
            padx=20,
            pady=8,
        ).pack(anchor="e", pady=(12, 0))

    def _open_module(self, module_id: str) -> None:
        if module_id == "wga":
            self.update_check_active = False
            for widget in self.root.winfo_children():
                widget.destroy()
            self.root.resizable(False, False)
            SplashScreen(self.root, initial_update_result=self.update_result)
            return

        self.update_check_active = False
        for widget in self.root.winfo_children():
            widget.destroy()
        self.root.overrideredirect(False)
        self.root.resizable(True, True)
        if module_id == "optimization":
            OptimizationUI(self.root, on_back=self._return_to_launcher)
        elif module_id == "network":
            NetworkMonitoringUI(self.root, on_back=self._return_to_launcher)

    def _return_to_launcher(self) -> None:
        for widget in self.root.winfo_children():
            widget.destroy()
        self.root.resizable(False, False)
        ProductLauncher(self.root)


# Началният loading екран на приложението.
class SplashScreen:
    # Помощна функция за init  .
    def __init__(self, root: tk.Tk, initial_update_result: UpdateResult | None = None) -> None:
        # Подготвя splash екрана и стартира анимацията.
        self.root = root
        self.root.title(APP_TITLE)
        apply_app_icon(self.root)
        self.dpi_value = apply_tk_dpi_scaling(self.root)
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        self.splash_width, self.splash_height = responsive_window_size(
            self.screen_width,
            self.screen_height,
            520,
            170,
        )
        self.splash_scale = self.splash_width / 520
        center_geometry(self.root, self.splash_width, self.splash_height)
        self.transparent_key = "#010203"
        self.root.configure(bg=self.transparent_key)
        self.root.resizable(False, False)
        self.root.overrideredirect(True)
        try:
            self.root.wm_attributes("-transparentcolor", self.transparent_key)
        except tk.TclError:
            pass

        self.progress_value = 0.0
        self.target_value = 0.0
        self.status_text = tk.StringVar(value="Стартиране на WGA...")
        self.message_queue: queue.Queue[tuple[str, float | str]] = queue.Queue()
        self.preloaded_state: dict[str, object] = {}
        if isinstance(initial_update_result, UpdateResult):
            self.preloaded_state["update_result"] = initial_update_result
        self.version_info = load_version_info()
        self.splash_active = True
        self.poll_job: str | None = None
        self.animation_job: str | None = None
        self.dashboard_job: str | None = None

        self.canvas = tk.Canvas(
            self.root,
            width=self.splash_width,
            height=self.splash_height,
            highlightthickness=0,
            bd=0,
            bg=self.transparent_key,
        )
        self.canvas.pack(fill="both", expand=True)

        self._draw_background()
        self._create_loader()
        self._start_boot_sequence()

    # Помощна функция за draw background.
    def _draw_background(self) -> None:
        # Този preloader е без тежък фон - оставяме само текст и бар.
        return

    # Създава loader и връща резултата към приложението.
    def _create_loader(self) -> None:
        # Създава минималистичен preloader с име и текст какво се зарежда.
        center_x = self.splash_width / 2
        scale = self.splash_scale
        self.canvas.create_text(
            center_x,
            48 * scale,
            text="WGA",
            fill="#effff5",
            font=("Segoe UI Semibold", max(26, int(34 * scale))),
        )

        self.status_label = self.canvas.create_text(
            center_x,
            84 * scale,
            text=self.status_text.get(),
            fill="#8df6b3",
            font=("Segoe UI", max(10, int(11 * scale))),
        )

        self.bar_width = 310 * scale
        self.bar_left = center_x - self.bar_width / 2
        self.bar_top = 108 * scale
        self.bar_height = max(18, 24 * scale)
        self.bar_radius = max(10, 12 * scale)

        self._draw_rounded_rect(
            self.bar_left,
            self.bar_top,
            self.bar_left + self.bar_width,
            self.bar_top + self.bar_height,
            self.bar_radius,
            fill="#d8d8d8",
            outline="",
        )
        self.progress_fill = self._draw_rounded_rect(
            self.bar_left,
            self.bar_top,
            self.bar_left + 4,
            self.bar_top + self.bar_height,
            self.bar_radius,
            fill="#14ff00",
            outline="",
        )
        self.progress_label = self.canvas.create_text(
            center_x,
            self.bar_top + self.bar_height / 2,
            text="0%",
            fill="#111111",
            font=("Segoe UI Semibold", max(11, int(14 * scale))),
        )

    # Помощна функция за draw rounded rect.
    def _draw_rounded_rect(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        radius: float,
        **kwargs: object,
    ) -> int:
        return self.canvas.create_polygon(
            self._rounded_rect_points(x1, y1, x2, y2, radius),
            smooth=True,
            splinesteps=30,
            **kwargs,
        )

    # Помощна функция за canvas alive.
    def _canvas_alive(self) -> bool:
        try:
            return bool(self.canvas.winfo_exists())
        except tk.TclError:
            return False

    # Помощна функция за rounded rect points.
    def _rounded_rect_points(self, x1: float, y1: float, x2: float, y2: float, radius: float) -> list[float]:
        safe_radius = min(radius, max(1.0, (x2 - x1) / 2), max(1.0, (y2 - y1) / 2))
        return [
            x1 + safe_radius,
            y1,
            x2 - safe_radius,
            y1,
            x2,
            y1,
            x2,
            y1 + safe_radius,
            x2,
            y2 - safe_radius,
            x2,
            y2,
            x2 - safe_radius,
            y2,
            x1 + safe_radius,
            y2,
            x1,
            y2,
            x1,
            y2 - safe_radius,
            x1,
            y1 + safe_radius,
            x1,
            y1,
        ]

    # Помощна функция за start boot sequence.
    def _start_boot_sequence(self) -> None:
        threading.Thread(target=self._run_startup_tasks, daemon=True).start()
        self._poll_queue()
        self._animate_progress()

    # Стартира startup tasks и връща резултата.
    def _run_startup_tasks(self) -> None:
        tasks = [
            ("Зареждане на основната конфигурация...", 0.12, self._load_configuration),
            ("Проверка на езиковите настройки...", 0.24, self._preload_language_status),
            ("Събиране на системната информация...", 0.42, self._preload_system_health),
            ("Проверка на състоянието на компонентите...", 0.60, self._preload_component_status),
            ("Проверка на наличния софтуер...", 0.80, self._preload_auto_installer_status),
            ("Проверка за актуализация...", 0.92, self._preload_update_status),
            ("Подготовка на интерфейса...", 1.00, self._finalize_startup),
        ]
        for label, progress, action in tasks:
            self.message_queue.put(("status", label))
            action()
            self.message_queue.put(("progress", progress))
        self.message_queue.put(("done", "Ready"))

    # Зарежда configuration от файл или конфигурация.
    def _load_configuration(self) -> None:
        # Зарежда базовите файлове, които после веднага трябват на UI-то.
        load_secure_store()
        load_settings()
        config_file = PROJECT_ROOT / "tasks.json"
        if config_file.exists():
            config_file.read_text(encoding="utf-8")

    # Помощна функция за preload language status.
    def _preload_language_status(self) -> None:
        # Взима езиковия статус предварително, за да не чака UI-то след старта.
        try:
            self.preloaded_state["language_status"] = get_language_status()
        except Exception as exc:
            self.preloaded_state["language_status_error"] = str(exc)

    # Помощна функция за preload system health.
    def _preload_system_health(self) -> None:
        # Взима живите системни данни предварително за dashboard-а.
        try:
            self.preloaded_state["health_items"] = collect_health_items()
        except Exception as exc:
            self.preloaded_state["health_error"] = str(exc)

    # Помощна функция за preload update status.
    def _preload_update_status(self) -> None:
        # Прави онлайн проверката предварително, за да няма второ мислене след старта.
        if isinstance(self.preloaded_state.get("update_result"), UpdateResult):
            return
        try:
            result = check_for_updates(
                str(self.version_info["version"]),
                str(self.version_info.get("version_info_url", "")),
            )
            self.preloaded_state["update_result"] = result
        except Exception as exc:
            self.preloaded_state["update_error"] = str(exc)

    # Подготвя dashboard probe според избраните настройки.
    def _build_dashboard_probe(self) -> "MainMenuUI":
        # Прави лек помощен обект, който ползва същите проверки без да строи целия UI.
        probe = MainMenuUI.__new__(MainMenuUI)
        probe.root = self.root
        probe.settings = load_settings()
        probe.secure_store = load_secure_store()
        probe.launch_info = get_launch_location_info()
        probe.resource_status = check_resource_status(PROJECT_ROOT)
        probe.version_info = self.version_info
        probe.office_inventory_cache = {}
        probe.office_online_cache = {}
        probe.office_maintenance_cache = {}
        probe.adobe_reader_status_cache = None
        probe.language_status_cache = self.preloaded_state.get("language_status")
        probe.nexus_admin_status_cache = None
        probe.program_selector_tasks_cache = []
        probe.program_selector_status_cache = {}
        probe.component_status_cache = None
        probe.latest_health_items = list(self.preloaded_state.get("health_items", [])) if isinstance(self.preloaded_state.get("health_items"), list) else []
        return probe

    # Помощна функция за preload component status.
    def _preload_component_status(self) -> None:
        # Подготвя десния панел със статусите на компонентите още преди UI-то.
        try:
            probe = self._build_dashboard_probe()
            self.preloaded_state["component_status_rows"] = probe._dashboard_component_rows()
        except Exception as exc:
            self.preloaded_state["component_status_error"] = str(exc)

    # Помощна функция за preload auto installer status.
    def _preload_auto_installer_status(self) -> None:
        # Подготвя задачите и статуса им за dashboard инсталатора още в preloader-а.
        try:
            probe = self._build_dashboard_probe()
            tasks = probe._auto_install_tasks()
            status_map: dict[str, tuple[bool, str]] = {}
            for task in tasks:
                status_map[task["id"]] = probe._safe_task_install_state(task)
            self.preloaded_state["program_selector_tasks"] = tasks
            self.preloaded_state["program_selector_status"] = status_map
        except Exception as exc:
            self.preloaded_state["program_selector_error"] = str(exc)

    # Помощна функция за finalize startup.
    def _finalize_startup(self) -> None:
        time.sleep(0.05)

    # Помощна функция за poll queue.
    def _poll_queue(self) -> None:
        if not self.splash_active or not self._canvas_alive():
            return
        try:
            while True:
                message_type, payload = self.message_queue.get_nowait()
                if message_type == "status":
                    self.status_text.set(str(payload))
                    if self._canvas_alive():
                        self.canvas.itemconfig(self.status_label, text=self.status_text.get())
                elif message_type == "progress":
                    self.target_value = float(payload)
                elif message_type == "done":
                    self.status_text.set("Всичко е заредено. Стартиране...")
                    if self._canvas_alive():
                        self.canvas.itemconfig(self.status_label, text=self.status_text.get())
                    self.target_value = 1.0
                    if self.dashboard_job is None:
                        self.dashboard_job = self.root.after(250, self._show_dashboard)
        except tk.TclError:
            self.splash_active = False
            return
        except queue.Empty:
            pass
        if self.splash_active and self._canvas_alive():
            self.poll_job = self.root.after(40, self._poll_queue)

    # Помощна функция за animate progress.
    def _animate_progress(self) -> None:
        if not self.splash_active or not self._canvas_alive():
            return
        if self.progress_value < self.target_value:
            delta = max(0.004, (self.target_value - self.progress_value) * 0.18)
            self.progress_value = min(self.target_value, self.progress_value + delta)
            self._update_bar()
        if self.splash_active and self._canvas_alive():
            self.animation_job = self.root.after(16, self._animate_progress)

    # Обновява bar след промяна в състоянието.
    def _update_bar(self) -> None:
        if not self._canvas_alive():
            return
        fill_width = max(4, self.bar_width * self.progress_value)
        points = self._rounded_rect_points(
            self.bar_left,
            self.bar_top,
            self.bar_left + fill_width,
            self.bar_top + self.bar_height,
            self.bar_radius,
        )
        try:
            self.canvas.coords(self.progress_fill, *points)
            self.canvas.itemconfig(self.progress_label, text=f"{int(self.progress_value * 100)}%")
        except tk.TclError:
            self.splash_active = False

    # Показва dashboard в интерфейса.
    def _show_dashboard(self) -> None:
        if not self.splash_active or not self._canvas_alive():
            return
        self.splash_active = False
        self.status_text.set("Изграждане на интерфейса...")
        try:
            self.canvas.itemconfig(self.status_label, text=self.status_text.get())
        except tk.TclError:
            return
        self.root.update_idletasks()
        for job_attr in ("poll_job", "animation_job", "dashboard_job"):
            job_id = getattr(self, job_attr)
            if job_id is not None:
                try:
                    self.root.after_cancel(job_id)
                except tk.TclError:
                    pass
                setattr(self, job_attr, None)
        MainMenuUI(
            self.root,
            startup_menu=get_startup_menu_from_args(),
            preloaded_state=self.preloaded_state,
        )
        if self._canvas_alive():
            try:
                self.canvas.destroy()
            except tk.TclError:
                pass
        try:
            self.root.overrideredirect(False)
        except tk.TclError:
            pass
        try:
            self.root.wm_attributes("-transparentcolor", "")
        except tk.TclError:
            pass
        self.root.resizable(True, True)
        apply_main_window_layout(self.root)


# Основният интерфейс след зареждане на splash екрана.
class MainMenuUI:
    # Помощна функция за init  .
    def __init__(self, root: tk.Tk, startup_menu: str | None = None, preloaded_state: dict[str, object] | None = None) -> None:
        # Тук пазим почти всички състояния, кешове и UI променливи.
        self.root = root
        apply_tk_dpi_scaling(self.root)
        apply_main_window_layout(self.root)
        self.root.configure(bg="#08130a")
        self.settings = load_settings()
        self.secure_store = load_secure_store()
        self.launch_info = get_launch_location_info()
        self.resource_status: ResourceStatus = check_resource_status(PROJECT_ROOT)
        self.version_info = load_version_info()
        self.dashboard_icons = self._load_dashboard_icon_sheet()
        self.menu_icons = self._load_menu_icons()
        self.dashboard_logo_large, self.dashboard_logo_small = self._load_dashboard_logo()
        self.activation_window: tk.Toplevel | None = None
        self.activation_status_var: tk.StringVar | None = None
        self.activation_progress_var: tk.IntVar | None = None
        self.activation_log_widget: tk.Text | None = None
        self.activation_close_button: tk.Button | None = None
        self.health_rows: list[tuple[tk.Label, tk.Label, tk.Label]] = []
        self.latest_health_items: list[HealthItem] = []
        self.health_refresh_job: str | None = None
        self.health_refresh_in_progress = False
        self.health_refresh_interval_ms = 2500
        self.health_canvas: tk.Canvas | None = None
        self.health_scrollbar: ttk.Scrollbar | None = None
        self.health_inner_frame: tk.Frame | None = None
        self.health_scroll_position = 0.0
        self.health_scroll_job: str | None = None
        self.dashboard_info_scroll_job: str | None = None
        self.dashboard_info_scroll_position = 0.0
        self.dashboard_render_job: str | None = None
        self.dashboard_host_frame: tk.Frame | None = None
        self.dashboard_is_rendering = False
        self.dashboard_live_widgets: dict[str, object] = {}
        self.update_result: UpdateResult | None = None
        self.update_download_url = ""
        self.update_package_url = ""
        self.update_installing = False
        self.update_popup_shown = False
        self.auto_install_vars: dict[str, tk.BooleanVar] = {}
        self.auto_remove_vars: dict[str, tk.BooleanVar] = {}
        self.auto_install_running = False
        self.program_selector_window: tk.Toplevel | None = None
        self.program_selector_tasks_cache: list[dict[str, str]] = []
        self.program_selector_status_cache: dict[str, tuple[bool, str]] = {}
        self.program_selector_scan_running = False
        self.component_status_cache: list[tuple[str, str, bool]] | None = None
        self.component_status_refresh_in_progress = False
        self.office_inventory_cache: dict[str, object] = {}
        self.office_online_cache: dict[str, object] = {}
        self.office_maintenance_cache: dict[str, object] = {}
        self.adobe_reader_status_cache: object | None = None
        self.language_status_cache: object | None = None
        self.language_status_var = tk.StringVar(value="Езиков статус: проверява се...")
        self.nexus_admin_status_cache: object | None = None

        self.history: list[str] = []
        self.current_menu = "main"
        self.current_page = 0
        self.startup_menu = startup_menu if startup_menu in MENU_TREE else "main"
        self.resize_render_job: str | None = None
        self.last_layout_bucket: tuple[int, int] = (-1, -1)
        self.ui_scale = 1.0
        self.sidebar_width = 320
        self.header_height_px = 90
        self.header_title_size = 22
        self.header_subtitle_size = 10
        self.body_text_size = 9
        self.button_text_size = 10
        self.section_title_size = 15
        self.card_title_size = 12
        self.card_desc_size = 9
        self.card_title_wrap = 320
        self.card_desc_wrap = 320
        self.compact_card_title_wrap = 350
        self.compact_card_desc_wrap = 350
        self.language_panel_width = 280
        self.language_status_wrap = 230
        self.system_info_wrap = 280
        self.resource_wrap = 520
        self.right_subtitle_wrap = 630
        self.content_pad_x = 20
        self.content_pad_y = 18
        self.nav_button_char_width = NAV_BUTTON_WIDTH
        self.card_button_width_px = CARD_BUTTON_PIXEL_WIDTH
        self.card_button_height_px = CARD_BUTTON_PIXEL_HEIGHT
        self.card_action_gap_px = 8
        self.scaled_card_min_height = CARD_MIN_HEIGHT
        self.scaled_menu_card_min_height = dict(MENU_CARD_MIN_HEIGHT)

        self.container = tk.Frame(self.root, bg=APP_BG)
        self.container.pack(fill="both", expand=True)

        self.header = tk.Frame(self.container, bg=APP_PANEL, height=96, bd=0, highlightthickness=1, highlightbackground=APP_BORDER)
        self.header.pack(fill="x")
        self.header.pack_propagate(False)

        self.title_label = tk.Label(
            self.header,
            text=APP_TITLE,
            font=("Segoe UI Semibold", 22),
            fg=APP_TEXT,
            bg=APP_PANEL,
        )
        self.title_label.pack(anchor="w", padx=26, pady=(12, 0))

        self.header_exit_button = tk.Button(
            self.header,
            text="Изход",
            command=self.root.destroy,
            font=("Segoe UI Semibold", 10),
            bg="#5a1d24",
            fg="#fff5f6",
            activebackground="#7f2831",
            activeforeground="#ffffff",
            bd=0,
            padx=18,
            pady=8,
            width=10,
            cursor="hand2",
        )
        self.header_exit_button.place(relx=1.0, x=-24, y=22, anchor="ne")

        self.header_dashboard_button = tk.Button(
            self.header,
            text="История на актуализациите",
            command=self._show_update_history,
            font=("Segoe UI Semibold", 10),
            bg=APP_ACCENT_SOFT,
            fg="#f2fff8",
            activebackground="#27a67a",
            activeforeground="#ffffff",
            bd=0,
            padx=18,
            pady=8,
            width=22,
            cursor="hand2",
        )

        self.subtitle_label = tk.Label(
            self.header,
            text="",
            font=("Segoe UI", 10),
            fg=APP_TEXT_SOFT,
            bg=APP_PANEL,
        )
        self.subtitle_label.pack(anchor="w", padx=26)

        self.header_device_chip = tk.Label(
            self.header,
            text=self._build_header_device_text(),
            font=("Segoe UI Semibold", 9),
            fg="#ddfff4",
            bg=APP_PANEL_ALT,
            padx=12,
            pady=5,
        )
        self.header_device_chip.place(x=26, y=56)

        self.version_chip = tk.Label(
            self.header,
            text=f"v{self.version_info['version']}",
            font=("Segoe UI Semibold", 9),
            fg="#effff8",
            bg=APP_ACCENT_SOFT,
            padx=10,
            pady=4,
        )
        self.version_chip.place(x=520, y=22)

        self.header_admin_chip = tk.Label(
            self.header,
            text="ADMIN MODE",
            font=("Segoe UI Semibold", 9),
            fg="#06110f",
            bg=APP_ACCENT,
            padx=10,
            pady=4,
        )
        self.header_admin_chip.place(x=608, y=22)

        self.content = tk.Frame(self.container, bg=APP_BG)
        self.content.pack(fill="both", expand=True, padx=20, pady=18)

        self.left_panel = tk.Frame(self.content, bg=APP_PANEL, width=320, bd=0, highlightthickness=1, highlightbackground=APP_BORDER)
        self.left_panel.pack(side="left", fill="y")
        self.left_panel.pack_propagate(False)
        self.sidebar_brand = tk.Frame(self.left_panel, bg=APP_PANEL)
        self.sidebar_brand.pack(fill="x", padx=14, pady=(10, 4))
        if self.dashboard_logo_large is not None:
            self.sidebar_brand_logo = tk.Label(self.sidebar_brand, image=self.dashboard_logo_large, bg=APP_PANEL)
        else:
            self.sidebar_brand_logo = tk.Label(self.sidebar_brand, text="🛡", font=("Segoe UI Symbol", 28), fg=APP_ACCENT, bg=APP_PANEL)
        self.sidebar_brand_logo.pack(anchor="center", pady=(0, 3))
        brand_text = tk.Frame(self.sidebar_brand, bg=APP_PANEL)
        brand_text.pack(fill="x", expand=True)
        tk.Label(
            brand_text,
            text="WinSys Guardian",
            font=("Segoe UI Semibold", 18),
            fg="#f4fff8",
            bg=APP_PANEL,
        ).pack(anchor="center", pady=(0, 0))
        tk.Label(
            brand_text,
            text="A D V A N C E D",
            font=("Segoe UI Semibold", 9),
            fg=APP_ACCENT,
            bg=APP_PANEL,
        ).pack(anchor="center", pady=(0, 0))
        self.sidebar_toggle_label = tk.Label(
            self.left_panel,
            text="☰",
            font=("Segoe UI Symbol", 18),
            fg=APP_TEXT_SOFT,
            bg=APP_PANEL,
        )
        self.sidebar_toggle_label.pack(anchor="e", padx=22, pady=(0, 4))

        self.menu_title = tk.Label(
            self.left_panel,
            text="Control Panel",
            font=("Segoe UI Semibold", 15),
            fg=APP_TEXT,
            bg=APP_PANEL,
        )

        self.menu_path = tk.Label(
            self.left_panel,
            text="Dashboard",
            justify="left",
            wraplength=280,
            font=("Segoe UI", 10),
            fg=APP_TEXT_SOFT,
            bg=APP_PANEL,
        )

        self.sidebar_nav_frame = tk.Frame(self.left_panel, bg=APP_PANEL)
        self.sidebar_nav_frame.pack(fill="x", padx=12, pady=(2, 8))
        self.sidebar_nav_buttons: dict[str, dict[str, tk.Widget]] = {}
        self.sidebar_section_label = tk.Label(
            self.sidebar_nav_frame,
            text="Navigation",
            font=("Segoe UI Semibold", 9),
            fg=APP_TEXT_MUTED,
            bg=APP_PANEL,
        )
        self.sidebar_section_label.pack(anchor="w", padx=8, pady=(0, 4))
        sidebar_text_map = {
            "main": "Dashboard\nОбзор",
            "activation": "Активация\nWindows и Office",
            "install_software": "Софтуер\nИнсталации",
            "auto_installer": "Auto Installer\nАвтоматични задачи",
            "driver_backup": "Архивиране\nДрайвери и отчет",
            "language": "Езици\nКлавиатури",
            "nexus_admin": "Nexus Admin\nAdmin инструменти",
        }
        sidebar_icon_map = {
            "main": "home_small",
            "activation": "key_small",
            "install_software": "download_small",
            "auto_installer": "robot_small",
            "language": "globe_small",
            "driver_backup": "drive_small",
            "nexus_admin": "admin_small",
        }
        for menu_key, label in SIDEBAR_SECTIONS:
            title, subtitle = (sidebar_text_map.get(menu_key, label).split("\n", 1) + [""])[:2]
            sidebar_image = self.menu_icons.get(f"{menu_key}_small") or self.dashboard_icons.get(sidebar_icon_map.get(menu_key, ""))
            card = tk.Frame(
                self.sidebar_nav_frame,
                bg=APP_PANEL_ALT,
                bd=0,
                highlightthickness=1,
                highlightbackground=APP_BORDER,
                cursor="hand2",
            )
            card.pack(fill="x", pady=2)
            stripe = tk.Frame(card, bg=APP_PANEL_ALT, width=4, cursor="hand2")
            stripe.pack(side="left", fill="y")
            body = tk.Frame(card, bg=APP_PANEL_ALT, cursor="hand2")
            body.pack(side="left", fill="both", expand=True, padx=8, pady=5)
            icon_label = tk.Label(
                body,
                image=sidebar_image,
                bg=APP_PANEL_ALT,
                cursor="hand2",
            )
            icon_label.pack(side="left", padx=(0, 8))
            text_box = tk.Frame(body, bg=APP_PANEL_ALT, cursor="hand2")
            text_box.pack(side="left", fill="both", expand=True)
            title_label = tk.Label(
                text_box,
                text=title,
                font=("Segoe UI Semibold", 9),
                fg=APP_TEXT,
                bg=APP_PANEL_ALT,
                anchor="w",
                cursor="hand2",
            )
            title_label.pack(anchor="w")
            subtitle_label = tk.Label(
                text_box,
                text=subtitle,
                font=("Segoe UI", 7),
                fg=APP_TEXT_SOFT,
                bg=APP_PANEL_ALT,
                justify="left",
                anchor="w",
                wraplength=180,
                cursor="hand2",
            )
            subtitle_label.pack(anchor="w", pady=(0, 0))
            arrow_label = tk.Label(
                body,
                text="›",
                font=("Segoe UI Semibold", 13),
                fg=APP_TEXT_MUTED,
                bg=APP_PANEL_ALT,
                cursor="hand2",
            )
            arrow_label.pack(side="right")
            for widget in (card, stripe, body, icon_label, text_box, title_label, subtitle_label, arrow_label):
                widget.bind("<Button-1>", lambda _event, key=menu_key: self._open_sidebar_menu(key))
            self.sidebar_nav_buttons[menu_key] = {
                "card": card,
                "stripe": stripe,
                "body": body,
                "icon": icon_label,
                "title": title_label,
                "subtitle": subtitle_label,
                "arrow": arrow_label,
            }
        self.system_info = tk.Label(
            self.left_panel,
            text=self._build_system_summary(),
            justify="left",
            wraplength=280,
            font=("Consolas", 9),
            fg="#d8fff2",
            bg=APP_PANEL_ALT,
            padx=12,
            pady=12,
        )

        self.hint_label = tk.Label(
            self.left_panel,
            text="Избери карта отдясно, за да отвориш секция или да стартираш подготвено действие.",
            wraplength=280,
            justify="left",
            font=("Segoe UI", 9),
            fg=APP_TEXT_MUTED,
            bg=APP_PANEL,
        )

        self.health_title = tk.Label(
            self.left_panel,
            text="Състояние на системата",
            font=("Segoe UI Semibold", 14),
            fg=APP_TEXT,
            bg=APP_PANEL,
        )

        self.health_frame = tk.Frame(
            self.left_panel,
            bg=APP_PANEL_ALT,
            bd=0,
            highlightthickness=1,
            highlightbackground=APP_BORDER,
        )

        self.health_loading_label = tk.Label(
            self.health_frame,
            text="Loading hardware diagnostics...",
            font=("Segoe UI", 10),
            fg=APP_TEXT_SOFT,
            bg=APP_PANEL_ALT,
            justify="left",
            wraplength=260,
        )
        self.health_loading_label.pack(anchor="w", padx=12, pady=12)
        self.sidebar_clock_var = tk.StringVar(value="--:--:--")
        self.sidebar_date_var = tk.StringVar(value="Зареждане...")
        self.sidebar_clock_card = tk.Frame(
            self.left_panel,
            bg="#0d1a14",
            bd=0,
            highlightthickness=1,
            highlightbackground=APP_BORDER,
        )
        self.sidebar_clock_card.pack(fill="x", padx=12, pady=(10, 10))
        tk.Label(
            self.sidebar_clock_card,
            textvariable=self.sidebar_clock_var,
            font=("Segoe UI Semibold", 16),
            fg=APP_ACCENT,
            bg="#0d1a14",
        ).pack(anchor="w", padx=14, pady=(10, 0))
        tk.Label(
            self.sidebar_clock_card,
            textvariable=self.sidebar_date_var,
            font=("Segoe UI", 9),
            fg=APP_TEXT_SOFT,
            bg="#0d1a14",
        ).pack(anchor="w", padx=14, pady=(0, 10))
        self.sidebar_version_card = tk.Frame(
            self.left_panel,
            bg="#0d1715",
            bd=0,
            highlightthickness=1,
            highlightbackground=APP_BORDER,
        )
        self.sidebar_version_card.pack(fill="x", padx=12, pady=(0, 14))
        version_header = tk.Frame(self.sidebar_version_card, bg="#0d1715")
        version_header.pack(fill="x", padx=12, pady=(10, 6))
        tk.Label(
            version_header,
            text=f"Версия: {self.version_info['version']}",
            font=("Segoe UI", 10),
            fg=APP_TEXT_SOFT,
            bg="#0d1715",
        ).pack(side="left")
        tk.Label(
            version_header,
            text="Най-нова",
            font=("Segoe UI Semibold", 8),
            fg=APP_ACCENT,
            bg="#113222",
            padx=8,
            pady=3,
        ).pack(side="right")
        tk.Label(
            self.sidebar_version_card,
            text="В© 2026 WinSys Guardian Team",
            font=("Segoe UI", 8),
            fg=APP_TEXT_MUTED,
            bg="#0d1715",
        ).pack(anchor="w", padx=12, pady=(0, 10))

        self.status_var = tk.StringVar(
            value=(
                f"Started from {self.launch_info['device_name']} "
                f"[{self.launch_info['drive']}] - {self.launch_info['drive_type_label']}"
            )
        )
        self.status_bar = tk.Label(
            self.container,
            textvariable=self.status_var,
            anchor="w",
            font=("Segoe UI", 10),
            fg="#d9fff3",
            bg=APP_PANEL,
            padx=18,
        )
        self.status_bar.pack(fill="x", side="bottom")

        self.right_panel = tk.Frame(self.content, bg=APP_BG)
        self.right_panel.pack(side="left", fill="both", expand=True, padx=(18, 0))

        self.card_title = tk.Label(
            self.right_panel,
            text="",
            font=("Segoe UI Semibold", 19),
            fg=APP_TEXT,
            bg=APP_BG,
        )
        self.card_title.pack(anchor="w")

        self.card_subtitle = tk.Label(
            self.right_panel,
            text="",
            font=("Segoe UI", 10),
            fg=APP_TEXT_SOFT,
            bg=APP_BG,
            wraplength=630,
            justify="left",
        )
        self.card_subtitle.pack(anchor="w", pady=(4, 12))

        self.overview_frame = tk.Frame(self.right_panel, bg=APP_BG)
        self.overview_frame.pack(fill="x", pady=(0, 12))
        self.overview_frame.columnconfigure(0, weight=1, uniform="overview")
        self.overview_frame.columnconfigure(1, weight=1, uniform="overview")
        self.overview_frame.columnconfigure(2, weight=1, uniform="overview")
        self.software_summary_frame = tk.Frame(
            self.right_panel,
            bg=APP_PANEL_ALT,
            bd=0,
            highlightthickness=1,
            highlightbackground=APP_BORDER,
        )

        self.overview_version_value = tk.StringVar(value=f"v{self.version_info['version']}")
        self.overview_resources_value = tk.StringVar(value="Ресурси: проверка...")
        self.overview_launch_value = tk.StringVar(value=self.launch_info["drive_type_label"])
        self.software_summary_resources_value = tk.StringVar(value=self._build_resource_summary())
        self.software_summary_mode_value = tk.StringVar(
            value=f"{self.launch_info['drive_type_label']} | {self.launch_info['drive']}"
        )
        self.overview_menu_value = tk.StringVar(value="Главно меню")

        self._build_overview_card(
            self.overview_frame,
            column=0,
            title="Версия",
            value_var=self.overview_version_value,
            subtitle="Текущ пакет на приложението",
            accent=APP_ACCENT,
            icon_name="shield_small",
        )
        self._build_overview_card(
            self.overview_frame,
            column=1,
            title="Инсталационни ресурси",
            value_var=self.overview_resources_value,
            subtitle="Наличност на локални и online пакети",
            accent=APP_ACCENT_BLUE,
            icon_name="download_small",
        )
        self._build_overview_card(
            self.overview_frame,
            column=2,
            title="Работен режим",
            value_var=self.overview_launch_value,
            subtitle="Откъде е стартирано приложението",
            accent=APP_WARNING,
            icon_name="drive_small",
        )

        self.update_banner = tk.Frame(
            self.right_panel,
            bg="#122229",
            bd=0,
            highlightthickness=1,
            highlightbackground="#1f4554",
        )
        self.update_banner.pack(fill="x", pady=(0, 12))

        self.update_icon_label = tk.Label(
            self.update_banner,
            text="i",
            font=self._font(16, "bold", "Segoe UI Semibold"),
            fg="#9de8ff",
            bg="#122229",
            width=2,
        )
        self.update_icon_label.pack(side="left", padx=(14, 8), pady=10)

        self.update_message_var = tk.StringVar(
            value=f"Проверка за актуализации за v{self.version_info['version']}..."
        )
        self.update_message_label = tk.Label(
            self.update_banner,
            textvariable=self.update_message_var,
            font=self._font(10),
            fg="#dcf8ff",
            bg="#122229",
            justify="left",
            anchor="w",
        )
        self.update_message_label.pack(side="left", fill="x", expand=True, pady=10)

        self.update_action_button = tk.Button(
            self.update_banner,
            text="Отвори",
            command=self._open_update_download,
            font=("Segoe UI Semibold", 9),
            bg="#1b5d73",
            fg="#f3fbff",
            activebackground="#267997",
            activeforeground="#ffffff",
            bd=0,
            padx=14,
            pady=7,
            state="disabled",
            cursor="hand2",
        )
        self.update_action_button.pack(side="right", padx=12, pady=8)

        self.update_history_button = tk.Button(
            self.update_banner,
            text="История",
            command=self._show_update_history,
            font=("Segoe UI Semibold", 9),
            bg="#173c4d",
            fg="#f3fbff",
            activebackground="#1b5d73",
            activeforeground="#ffffff",
            bd=0,
            padx=12,
            pady=7,
            cursor="hand2",
        )

        self.resource_frame = tk.Frame(
            self.right_panel,
            bg=APP_PANEL_ALT,
            bd=0,
            highlightthickness=1,
            highlightbackground=APP_BORDER,
        )
        self.resource_frame.pack(fill="x", pady=(0, 12))

        self.resource_title = tk.Label(
            self.resource_frame,
            text="\u0418\u043d\u0441\u0442\u0430\u043b\u0430\u0446\u0438\u043e\u043d\u043d\u0438 \u0440\u0435\u0441\u0443\u0440\u0441\u0438",
            font=("Segoe UI Semibold", 11),
            fg=APP_TEXT,
            bg=APP_PANEL_ALT,
        )
        self.resource_title.pack(side="left", padx=(14, 10), pady=10)

        self.resource_status_label = tk.Label(
            self.resource_frame,
            text=self._build_resource_summary(),
            justify="left",
            anchor="w",
            wraplength=520,
            font=("Segoe UI", 9),
            fg=self._resource_status_color(),
            bg=APP_PANEL_ALT,
        )
        self.resource_status_label.pack(side="left", fill="x", expand=True, pady=10)

        self.resource_download_button = tk.Button(
            self.resource_frame,
            text="\u0418\u0437\u0442\u0435\u0433\u043b\u0438",
            command=self._download_missing_resources,
            font=("Segoe UI Semibold", 9),
            bg="#78561c",
            fg="#fff7d6",
            activebackground="#9a722a",
            activeforeground="#ffffff",
            bd=0,
            padx=12,
            pady=7,
            cursor="hand2",
        )
        self.resource_download_button.pack(side="right", padx=(6, 12), pady=10)

        self.resource_details_button = tk.Button(
            self.resource_frame,
            text="\u0414\u0435\u0442\u0430\u0439\u043b\u0438",
            command=self._show_resource_details,
            font=("Segoe UI Semibold", 9),
            bg=APP_ACCENT_SOFT,
            fg="#eefef1",
            activebackground="#27a67a",
            activeforeground="#ffffff",
            bd=0,
            padx=12,
            pady=7,
            cursor="hand2",
        )
        self.resource_details_button.pack(side="right", pady=10)
        self._refresh_resource_panel()

        summary_left = tk.Frame(self.software_summary_frame, bg=APP_PANEL_ALT)
        summary_left.pack(side="left", fill="x", expand=True, padx=(12, 8), pady=8)
        tk.Label(
            summary_left,
            text="Инсталационни ресурси",
            font=("Segoe UI Semibold", 9),
            fg=APP_TEXT_SOFT,
            bg=APP_PANEL_ALT,
        ).pack(anchor="w")
        self.software_summary_resources_label = tk.Label(
            summary_left,
            textvariable=self.software_summary_resources_value,
            font=("Segoe UI Semibold", 10),
            fg=self._resource_status_color(),
            bg=APP_PANEL_ALT,
            anchor="w",
            justify="left",
            wraplength=520,
        )
        self.software_summary_resources_label.pack(anchor="w", fill="x", pady=(2, 0))

        summary_action = tk.Frame(self.software_summary_frame, bg=APP_PANEL_ALT, width=220, height=54)
        summary_action.pack(side="left", fill="y", padx=(0, 8), pady=8)
        summary_action.pack_propagate(False)
        self.software_summary_download_button = tk.Button(
            summary_action,
            text="Изтегли липсващите",
            command=self._download_missing_resources,
            font=("Segoe UI Semibold", 9),
            bg="#384039",
            fg="#9aa69c",
            activebackground="#8b7432",
            activeforeground="#ffffff",
            bd=0,
            padx=12,
            pady=7,
            cursor="hand2",
            state="disabled",
        )
        self.software_summary_download_button.place(relx=0.5, rely=0.5, anchor="center", width=190, height=36)

        summary_right = tk.Frame(self.software_summary_frame, bg=APP_PANEL_ALT)
        summary_right.pack(side="left", fill="x", expand=True, padx=(8, 12), pady=8)
        tk.Label(
            summary_right,
            text="Работен режим",
            font=("Segoe UI Semibold", 9),
            fg=APP_TEXT_SOFT,
            bg=APP_PANEL_ALT,
        ).pack(anchor="w")
        tk.Label(
            summary_right,
            textvariable=self.software_summary_mode_value,
            font=("Segoe UI Semibold", 10),
            fg=APP_WARNING,
            bg=APP_PANEL_ALT,
            anchor="w",
            justify="left",
            wraplength=360,
        ).pack(anchor="w", fill="x", pady=(2, 0))
        self._refresh_resource_panel()

        self.nav_frame = tk.Frame(
            self.right_panel,
            bg=APP_BG,
            bd=0,
            highlightthickness=1,
            highlightbackground=APP_BORDER,
        )
        self.nav_frame.pack(fill="x", side="bottom", pady=(10, 0))

        self.page_label = tk.Label(
            self.nav_frame,
            text="Page 1 / 1",
            font=("Segoe UI", 10),
            fg=APP_TEXT_SOFT,
            bg=APP_BG,
        )
        self.page_label.pack(side="left")

        self.controls_frame = tk.Frame(self.nav_frame, bg=APP_BG)
        self.controls_frame.pack(side="right")

        self.prev_button = self._make_nav_button(self.controls_frame, "\u041f\u0440\u0435\u0434\u0438\u0448\u043d\u0430", self.previous_page)
        self.prev_button.pack(side="left", padx=(0, 6))
        self.next_button = self._make_nav_button(self.controls_frame, "\u041d\u0430\u043f\u0440\u0435\u0434", self.next_page)
        self.next_button.pack(side="left", padx=(0, 6))
        self.back_button = self._make_nav_button(self.controls_frame, "\u041d\u0430\u0437\u0430\u0434", self.go_back, accent="#17361f")
        self.back_button.pack(side="left", padx=(0, 6))
        self.dashboard_button = self._make_nav_button(self.controls_frame, "Dashboard", self.go_dashboard, accent="#17361f")
        self.dashboard_button.pack(side="left", padx=(0, 6))
        self.exit_button = self._make_nav_button(self.controls_frame, "\u0418\u0437\u0445\u043e\u0434", self.root.destroy, accent="#7a1f1f")
        self.exit_button.pack(side="left")

        self.cards_frame = tk.Frame(self.right_panel, bg=APP_BG)
        self.cards_frame.pack(fill="both", expand=True)

        self.language_status_panel = tk.Frame(
            self.cards_frame,
            bg=APP_PANEL,
            width=280,
            bd=0,
            highlightthickness=1,
            highlightbackground=APP_BORDER,
        )
        self.language_status_panel.grid_propagate(False)

        self.language_status_title = tk.Label(
            self.language_status_panel,
            text="Език и клавиатури",
            font=("Segoe UI Semibold", 13),
            fg=APP_TEXT,
            bg=APP_PANEL,
        )
        self.language_status_title.pack(anchor="w", padx=14, pady=(14, 6))

        self.language_status_frame = tk.Frame(
            self.language_status_panel,
            bg=APP_PANEL_ALT,
            bd=0,
            highlightthickness=1,
            highlightbackground=APP_BORDER,
        )
        self.language_status_frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        self.language_status_label = tk.Label(
            self.language_status_frame,
            textvariable=self.language_status_var,
            justify="left",
            wraplength=230,
            font=("Segoe UI", 9),
            fg="#ddfff3",
            bg=APP_PANEL_ALT,
        )
        self.language_status_label.pack(anchor="w", fill="x", padx=12, pady=10)

        self._update_layout_metrics()
        self._apply_responsive_theme()
        self._apply_startup_preload(preloaded_state or {})
        self._refresh_overview_cards()
        self._update_sidebar_clock()
        self.render_menu(self.startup_menu, reset_history=True)
        self.root.bind("<Configure>", self._on_root_resize, add="+")
        self.root.bind_all("<Control-Shift-2>", self._show_hidden_menu, add="+")
        self.root.bind_all("<Control-KeyPress-2>", self._show_hidden_menu, add="+")
        self.root.bind_all("<Control-Shift-KeyPress-2>", self._show_hidden_menu, add="+")
        self.root.bind_all("<Control-Shift-KeyPress-@>", self._show_hidden_menu, add="+")
        self.root.bind_all("<Control-Shift-KeyPress-quotedbl>", self._show_hidden_menu, add="+")
        self.root.bind_all("<Control-Shift-KeyPress>", self._handle_ctrl_shift_keypress, add="+")

    # Показва скритото меню при клавишна команда.
    def _show_hidden_menu(self, event: tk.Event | None = None) -> None:
        if self.current_menu == "hidden_menu":
            return
        self.status_var.set("Скритото меню е активирано.")
        self.render_menu("hidden_menu", reset_history=True)

    def _handle_ctrl_shift_keypress(self, event: tk.Event) -> None:
        if event.state & 0x4 and event.state & 0x1:  # Ctrl and Shift pressed
            if event.keysym in {"2", "@", "quotedbl", "at", "numbersign", "sterling"}:
                self._show_hidden_menu(event)

    # Показва agent статуса от локален JSON файл.
    def _show_agent_status(self) -> None:
        if not AGENT_STATUS_FILE.exists():
            self.status_var.set("Agent status file не е намерен.")
            messagebox.showwarning(
                "Agent Status Missing",
                f"Файлът {AGENT_STATUS_FILE.name} не е намерен. Изпълнете wga_agent.py на машината и опитайте отново.",
                parent=self.root,
            )
            return

        try:
            agent_data = json.loads(AGENT_STATUS_FILE.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            self.status_var.set("Неуспешно четене на agent статусния файл.")
            messagebox.showerror(
                "Agent Status Error",
                "Файлът с агентния статус не може да бъде прочетен или е повреден.",
                parent=self.root,
            )
            return

        summary_lines = [
            f"Hostname: {agent_data.get('hostname', 'N/A')}",
            f"Platform: {agent_data.get('platform', 'N/A')} {agent_data.get('platform_release', '')}",
            f"Processor: {agent_data.get('processor', 'N/A')}",
            f"Python: {agent_data.get('python_version', 'N/A')}",
            f"Online: {'Yes' if agent_data.get('online') else 'No'}",
            f"Local IPs: {', '.join(agent_data.get('local_ips', [])) or 'N/A'}",
            f"Timestamp: {agent_data.get('timestamp', 'N/A')}",
        ]
        self.status_var.set("Agent статусът е зареден.")
        messagebox.showinfo(
            "Agent Status",
            "\n".join(summary_lines),
            parent=self.root,
        )

    # Подготвя system summary според избраните настройки.
    def _build_system_summary(self) -> str:
        return (
            f"OS: {platform.system()} {platform.release()}\n"
            f"CPU Threads: {os.cpu_count() or 'N/A'}\n"
            f"Device: {self.launch_info['device_name']}\n"
            f"Drive: {self.launch_info['drive']}\n"
            f"Drive Type: {self.launch_info['drive_type_label']}\n"
            f"Start Path: {self.launch_info['program_path']}\n"
            f"Installers Root: {self.launch_info['installers_root']}\n"
            f"Installers Available: {self.launch_info['installers_available']}\n"
            f"App Version: {self.version_info['version']}\n"
            "Mode: Portable Admin UI"
        )

    # Помощна функция за apply startup preload.
    def _apply_startup_preload(self, preloaded_state: dict[str, object]) -> None:
        # Вкарва вече заредените данни в UI-то, за да не чакаме повторно след preloader-а.
        language_status = preloaded_state.get("language_status")
        if isinstance(language_status, LanguageStatus):
            self.language_status_cache = language_status
            self._apply_language_status_summary(
                self._build_language_status_summary(language_status),
                "#9aff9f" if language_status.has_language_pack or language_status.has_bulgarian else "#ffb0a8",
            )
        else:
            self._load_language_status_async()

        health_items = preloaded_state.get("health_items")
        if isinstance(health_items, list):
            self.latest_health_items = [item for item in health_items if isinstance(item, HealthItem)]
        if self.latest_health_items:
            self._apply_system_health_update(self.latest_health_items)
        else:
            self._load_system_health_async()

        update_result = preloaded_state.get("update_result")
        if isinstance(update_result, UpdateResult):
            # Стартовият launcher вече е показал update статуса; не отваряме втори popup в WGA.
            self.update_popup_shown = True
            self._apply_update_result(update_result)
        else:
            self._check_updates_async()

        program_tasks = preloaded_state.get("program_selector_tasks")
        if isinstance(program_tasks, list):
            self.program_selector_tasks_cache = [dict(task) for task in program_tasks if isinstance(task, dict)]
        program_status = preloaded_state.get("program_selector_status")
        if isinstance(program_status, dict):
            self.program_selector_status_cache = {
                str(task_id): value
                for task_id, value in program_status.items()
                if isinstance(value, tuple) and len(value) == 2
            }

        component_rows = preloaded_state.get("component_status_rows")
        if isinstance(component_rows, list):
            self.component_status_cache = [
                tuple(row)
                for row in component_rows
                if isinstance(row, tuple) and len(row) == 3
            ]

    # Подготвя header device text според избраните настройки.
    def _build_header_device_text(self) -> str:
        # Кратък статус в header-а откъде е стартирано приложението.
        return (
            f"{self.launch_info['drive_type_label']}  |  "
            f"{self.launch_info['drive']}  |  "
            f"{self.launch_info['device_name']}"
        )

    # Зарежда dashboard icon sheet от файл или конфигурация.
    def _load_dashboard_icon_sheet(self) -> dict[str, tk.PhotoImage]:
        # Реже иконите от общия sprite sheet, за да ги ползваме по целия dashboard.
        png_icons = self._load_dashboard_png_icons()
        if png_icons:
            return png_icons

        sheet_path = runtime_file(DASHBOARD_ICON_SHEET_RELATIVE)
        if not sheet_path.exists():
            return {}

        try:
            sprite = tk.PhotoImage(file=str(sheet_path))
        except tk.TclError:
            return {}

        columns = 4
        rows = 4
        cell_width = max(1, sprite.width() // columns)
        cell_height = max(1, sprite.height() // rows)
        names = [
            "shield",
            "home",
            "key",
            "download",
            "robot",
            "globe",
            "drive",
            "admin",
            "refresh",
            "cpu",
            "bolt",
            "ram",
            "disk",
            "warning",
            "monitor",
            "actions",
        ]
        icons: dict[str, tk.PhotoImage] = {"sheet": sprite}
        for index, name in enumerate(names):
            col = index % columns
            row = index // columns
            left = col * cell_width
            top = row * cell_height
            right = left + cell_width
            bottom = top + cell_height
            cropped = tk.PhotoImage()
            cropped.tk.call(cropped, "copy", sprite, "-from", left, top, right, bottom)
            icons[name] = cropped.subsample(6, 6)
            icons[f"{name}_small"] = cropped.subsample(9, 9)
        return icons

    # Зарежда dashboard logo от файл или конфигурация.
    def _load_dashboard_png_icons(self) -> dict[str, tk.PhotoImage]:
        manifest_path = runtime_file(DASHBOARD_ICONS_MANIFEST_RELATIVE)
        if not manifest_path.exists():
            return {}
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
        except (OSError, json.JSONDecodeError):
            return {}

        icons: dict[str, tk.PhotoImage] = {}
        for icon_key, spec in manifest.items():
            if not isinstance(spec, dict):
                continue
            icon_file = spec.get("file")
            if not isinstance(icon_file, str):
                continue
            icon_path = runtime_file(icon_file)
            if not icon_path.exists():
                continue
            try:
                image = tk.PhotoImage(file=str(icon_path))
            except tk.TclError:
                continue
            base_size = max(1, min(image.width(), image.height()))
            divisor_large = max(1, base_size // 64)
            divisor_small = max(1, base_size // 34)
            icons[str(icon_key)] = image.subsample(divisor_large, divisor_large)
            icons[f"{icon_key}_small"] = image.subsample(divisor_small, divisor_small)
        return icons

    def _load_menu_icons(self) -> dict[str, tk.PhotoImage]:
        manifest_path = runtime_file(MENU_ICONS_MANIFEST_RELATIVE)
        if not manifest_path.exists():
            return {}
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return {}

        icons: dict[str, tk.PhotoImage] = {}
        for menu_key, spec in manifest.items():
            if not isinstance(spec, dict):
                continue
            icon_file = spec.get("file")
            if not isinstance(icon_file, str):
                continue
            icon_path = runtime_file(icon_file)
            if not icon_path.exists():
                continue
            try:
                image = tk.PhotoImage(file=str(icon_path))
            except tk.TclError:
                continue
            base_size = max(1, min(image.width(), image.height()))
            divisor_large = max(1, base_size // 64)
            divisor_card = max(1, base_size // 46)
            divisor_small = max(1, base_size // 30)
            icons[f"{menu_key}_large"] = image.subsample(divisor_large, divisor_large)
            icons[f"{menu_key}_card"] = image.subsample(divisor_card, divisor_card)
            icons[f"{menu_key}_small"] = image.subsample(divisor_small, divisor_small)
            icons[menu_key] = icons[f"{menu_key}_card"]
        return icons

    def _menu_icon_for_item(self, item: dict[str, str], size: str = "card") -> tk.PhotoImage | None:
        if item.get("kind") == "menu":
            menu_key = item.get("target", "")
            return self.menu_icons.get(f"{menu_key}_{size}") or self.menu_icons.get(menu_key)
        action_id = item.get("action_id", "")
        if action_id:
            action_key = f"action_{action_id}"
            action_icon = self.menu_icons.get(f"{action_key}_{size}") or self.menu_icons.get(action_key)
            if action_icon is not None:
                return action_icon
        label_key = f"card_{self.current_menu}_{self._menu_icon_key(item.get('label', ''))}"
        label_icon = self.menu_icons.get(f"{label_key}_{size}") or self.menu_icons.get(label_key)
        if label_icon is not None:
            return label_icon
        action_icon_map = {
            "add_desktop_icons": "main",
            "open_console": "nexus_admin",
            "driver_pc_report": "driver_backup",
        }
        mapped_key = action_icon_map.get(action_id)
        if mapped_key:
            return self.menu_icons.get(f"{mapped_key}_{size}") or self.menu_icons.get(mapped_key)
        return None

    def _menu_icon_key(self, value: str) -> str:
        normalized = value.lower().replace("&", "and")
        normalized = re.sub(r"[^a-z0-9]+", "_", normalized).strip("_")
        return normalized or "card"

    def _load_dashboard_logo(self) -> tuple[tk.PhotoImage | None, tk.PhotoImage | None]:
        # Зарежда голямо и малко лого за dashboard панелите.
        logo_path = runtime_file(APP_LOGO_RELATIVE)
        if not logo_path.exists():
            return None, None
        try:
            image = tk.PhotoImage(file=str(logo_path))
        except tk.TclError:
            return None, None
        return image.subsample(7, 7), image.subsample(11, 11)

    # Подготвя overview card според избраните настройки.
    def _build_overview_card(
        self,
        parent: tk.Widget,
        *,
        column: int,
        title: str,
        value_var: tk.StringVar,
        subtitle: str,
        accent: str,
        icon_name: str = "",
    ) -> None:
        # Малка dashboard карта в горната част на екрана.
        card = tk.Frame(
            parent,
            bg=APP_PANEL,
            bd=0,
            highlightthickness=1,
            highlightbackground=APP_BORDER,
        )
        card.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 6, 0), pady=0)
        tk.Frame(card, bg=accent, height=4).pack(fill="x")
        body = tk.Frame(card, bg=APP_PANEL)
        body.pack(fill="both", expand=True, padx=14, pady=12)
        icon_image = self.dashboard_icons.get(icon_name)
        if icon_image is not None:
            tk.Label(body, image=icon_image, bg=APP_PANEL).pack(anchor="w", pady=(0, 6))
        tk.Label(
            body,
            text=title,
            font=("Segoe UI", 9),
            fg=APP_TEXT_MUTED,
            bg=APP_PANEL,
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            body,
            textvariable=value_var,
            font=("Segoe UI Semibold", 14),
            fg=APP_TEXT,
            bg=APP_PANEL,
            anchor="w",
            justify="left",
            wraplength=240,
        ).pack(anchor="w", pady=(4, 2))
        tk.Label(
            body,
            text=subtitle,
            font=("Segoe UI", 9),
            fg=APP_TEXT_SOFT,
            bg=APP_PANEL,
            anchor="w",
            justify="left",
            wraplength=240,
        ).pack(anchor="w")

    # Помощна функция за sidebar section for menu.
    def _sidebar_section_for_menu(self, menu_key: str) -> str:
        # Свързва подменютата с основната секция в левия sidebar.
        activation_group = {"activation", "windows10_activation", "windows11_activation", "office_activation"}
        install_group = {"install_software", "office_install_center", "secret_install", "office_center"}
        if menu_key in activation_group:
            return "activation"
        if menu_key in install_group:
            return "install_software"
        if menu_key.startswith("online_") or menu_key.startswith("install_office_"):
            return "install_software"
        if menu_key in {"language"}:
            return "language"
        if menu_key in {"driver_backup"}:
            return "driver_backup"
        if menu_key in {"nexus_admin"}:
            return "nexus_admin"
        if menu_key in {"auto_installer"}:
            return "auto_installer"
        return "main"

    # Помощна функция за refresh sidebar navigation.
    def _refresh_sidebar_navigation(self) -> None:
        # Оцветява активната секция в sidebar-а.
        active_key = self._sidebar_section_for_menu(self.current_menu)
        for menu_key, parts in self.sidebar_nav_buttons.items():
            is_active = menu_key == active_key
            card_bg = "#103125" if is_active else APP_PANEL_ALT
            border = APP_BORDER_STRONG if is_active else APP_BORDER
            title_fg = "#effff7" if is_active else APP_TEXT
            subtitle_fg = "#c9f7e2" if is_active else APP_TEXT_SOFT
            arrow_fg = APP_ACCENT if is_active else APP_TEXT_MUTED
            stripe_bg = APP_ACCENT if is_active else APP_PANEL_ALT
            parts["card"].configure(bg=card_bg, highlightbackground=border)
            parts["stripe"].configure(bg=stripe_bg)
            parts["body"].configure(bg=card_bg)
            parts["icon"].configure(bg=card_bg)
            parts["title"].configure(bg=card_bg, fg=title_fg)
            parts["subtitle"].configure(bg=card_bg, fg=subtitle_fg)
            parts["arrow"].configure(bg=card_bg, fg=arrow_fg)

    # Отваря sidebar menu или съответния прозорец.
    def _open_sidebar_menu(self, menu_key: str) -> None:
        # Бърза навигация от лявото меню.
        if menu_key == "main":
            self.go_dashboard()
            return
        self.history = ["main"]
        self.render_menu(menu_key)

    # Помощна функция за refresh overview cards.
    def _refresh_overview_cards(self) -> None:
        # Обновява кратките статуси в горния dashboard ред.
        self.overview_version_value.set(f"v{self.version_info['version']}")
        if not self.resource_status.configured:
            resources_text = "Няма manifest"
        elif self.resource_status.complete:
            resources_text = f"Пълна готовност {self.resource_status.available}/{self.resource_status.total}"
        else:
            resources_text = f"Липсват {self.resource_status.missing} пакета"
        self.overview_resources_value.set(resources_text)
        self.overview_launch_value.set(
            f"{self.launch_info['drive_type_label']}\n{self.launch_info['drive']} • {self.launch_info['device_name']}"
        )
        self.overview_menu_value.set(MENU_TREE[self.current_menu]["title"])

    # Обновява sidebar clock след промяна в състоянието.
    def _update_sidebar_clock(self) -> None:
        # Поддържа часовника вляво като на mockup-а.
        weekday_names = [
            "понеделник",
            "вторник",
            "сряда",
            "четвъртък",
            "петък",
            "събота",
            "неделя",
        ]
        month_names = [
            "януари",
            "февруари",
            "март",
            "април",
            "май",
            "юни",
            "юли",
            "август",
            "септември",
            "октомври",
            "ноември",
            "декември",
        ]
        now = datetime.now()
        self.sidebar_clock_var.set(now.strftime("%H:%M:%S"))
        self.sidebar_date_var.set(
            f"{weekday_names[now.weekday()]}, {now.day} {month_names[now.month - 1]} {now.year}"
        )
        self.root.after(1000, self._update_sidebar_clock)

    # Помощна функция за is activation menu.
    def _is_activation_menu(self, menu_key: str) -> bool:
        return menu_key in {"activation", "windows10_activation", "windows11_activation", "office_activation"}

    # Помощна функция за toggle dashboard chrome.
    def _toggle_dashboard_chrome(
        self,
        dashboard_mode: bool,
        hide_overview: bool = False,
        hide_update_banner: bool = False,
        show_resource_panel: bool = False,
        show_software_summary: bool = False,
    ) -> None:
        # Скрива старите общи панели, когато сме на новия dashboard изглед.
        if dashboard_mode:
            for widget in (
                self.card_title,
                self.card_subtitle,
                self.overview_frame,
                self.update_banner,
                self.resource_frame,
                self.software_summary_frame,
                self.nav_frame,
            ):
                if widget.winfo_manager():
                    widget.pack_forget()
            if self.cards_frame.winfo_manager():
                self.cards_frame.pack_configure(fill="both", expand=True)
            else:
                self.cards_frame.pack(fill="both", expand=True)
            self.cards_frame.lift()
            return

        if not self.card_title.winfo_manager():
            self.card_title.pack(anchor="w", before=self.cards_frame)
        if not self.card_subtitle.winfo_manager():
            self.card_subtitle.pack(anchor="w", pady=(4, 12), before=self.cards_frame)
        if hide_overview:
            if self.overview_frame.winfo_manager():
                self.overview_frame.pack_forget()
        elif not self.overview_frame.winfo_manager():
            self.overview_frame.pack(fill="x", pady=(0, 12), before=self.cards_frame)
        if hide_update_banner:
            if self.update_banner.winfo_manager():
                self.update_banner.pack_forget()
        elif not self.update_banner.winfo_manager():
            self.update_banner.pack(fill="x", pady=(0, 12), before=self.cards_frame)
        if not show_resource_panel:
            if self.resource_frame.winfo_manager():
                self.resource_frame.pack_forget()
        elif not self.resource_frame.winfo_manager():
            self.resource_frame.pack(fill="x", pady=(0, 12), before=self.cards_frame)
        if not show_software_summary:
            if self.software_summary_frame.winfo_manager():
                self.software_summary_frame.pack_forget()
        elif not self.software_summary_frame.winfo_manager():
            self.software_summary_frame.pack(fill="x", pady=(0, 8), before=self.cards_frame)
        if not self.nav_frame.winfo_manager():
            self.nav_frame.pack(fill="x", side="bottom", pady=(10, 0))

    # Помощна функция за rounded rect points.
    def _rounded_rect_points(self, x1: int, y1: int, x2: int, y2: int, radius: int) -> list[int]:
        # Смята точките за заоблен правоъгълник върху Canvas.
        r = max(4, min(radius, (x2 - x1) // 2, (y2 - y1) // 2))
        return [
            x1 + r, y1,
            x2 - r, y1,
            x2, y1,
            x2, y1 + r,
            x2, y2 - r,
            x2, y2,
            x2 - r, y2,
            x1 + r, y2,
            x1, y2,
            x1, y2 - r,
            x1, y1 + r,
            x1, y1,
        ]

    # Подготвя soft panel според избраните настройки.
    def _build_soft_panel(
        self,
        parent: tk.Widget,
        *,
        panel_bg: str,
        border: str,
        radius: int = 18,
        base_bg: str | None = None,
    ) -> tk.Frame:
        # Прави panel с по-меки, заоблени ъгли за dashboard секциите.
        host_bg = base_bg or str(parent.cget("bg"))
        outer = tk.Frame(parent, bg=host_bg, bd=0, highlightthickness=0)
        inner = tk.Frame(outer, bg=panel_bg, bd=0, highlightthickness=0)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        canvas = tk.Canvas(
            outer,
            bg=host_bg,
            bd=0,
            highlightthickness=0,
            relief="flat",
        )
        canvas.place(relx=0, rely=0, relwidth=1, relheight=1)
        inner.lift()

        # Помощна функция за redraw.
        def redraw(_event: tk.Event | None = None) -> None:
            width = max(outer.winfo_width() - 2, 40)
            height = max(outer.winfo_height() - 2, 40)
            canvas.delete("shape")
            points = self._rounded_rect_points(1, 1, width, height, radius)
            canvas.create_polygon(
                points,
                smooth=True,
                splinesteps=36,
                fill=panel_bg,
                outline=border,
                width=1,
                tags="shape",
            )

        outer.bind("<Configure>", redraw)
        outer.content = inner  # type: ignore[attr-defined]
        outer.redraw_panel = redraw  # type: ignore[attr-defined]
        return outer

    # Помощна функция за resource status color.
    def _resource_status_color(self) -> str:
        if not self.resource_status.configured:
            return "#f9e6a8"
        if self.resource_status.complete:
            return "#9aff9f"
        if self.resource_status.downloadable_missing:
            return "#ffe08a"
        return "#ffb0a8"

    # Подготвя resource summary според избраните настройки.
    def _build_resource_summary(self) -> str:
        status = self.resource_status
        if not status.configured:
            return "Manifest не е намерен. Няма списък с нужните инсталационни файлове."
        if status.complete:
            state = "[OK] \u0412\u0441\u0438\u0447\u043a\u043e \u0435 \u043d\u0430\u043b\u0438\u0447\u043d\u043e"
        else:
            state = f"[\u041b\u0418\u041f\u0421\u0418] \u041b\u0438\u043f\u0441\u0432\u0430\u0442 {status.missing} \u043f\u0430\u043a\u0435\u0442\u0430"
        return (
            f"{state} | \u041d\u0430\u043b\u0438\u0447\u043d\u0438: {status.available}/{status.total} | "
            f"\u0417\u0430 \u0438\u0437\u0442\u0435\u0433\u043b\u044f\u043d\u0435: {status.downloadable_missing}\n"
            f"\u041d\u043e\u0441\u0438\u0442\u0435\u043b: {self.launch_info['drive_type_label']} "
            f"{self.launch_info['drive']} | Installers: {status.installers_root}"
        )

    # Помощна функция за refresh resource panel.
    def _refresh_resource_panel(self) -> None:
        self.launch_info = get_launch_location_info()
        self.resource_status = check_resource_status(PROJECT_ROOT)
        if hasattr(self, "system_info"):
            self.system_info.config(text=self._build_system_summary())
        if hasattr(self, "header_device_chip"):
            self.header_device_chip.config(text=self._build_header_device_text())
        if hasattr(self, "overview_resources_value"):
            self._refresh_overview_cards()
        self.resource_status_label.config(
            text=self._build_resource_summary(),
            fg=self._resource_status_color(),
        )
        if hasattr(self, "software_summary_resources_value"):
            self.software_summary_resources_value.set(self._build_resource_summary())
        if hasattr(self, "software_summary_resources_label"):
            self.software_summary_resources_label.config(fg=self._resource_status_color())
        can_download = self.resource_status.missing > 0
        download_button_state = {
            "state": "normal" if can_download else "disabled",
            "bg": "#7d6a2d" if can_download else "#384039",
            "fg": "#fff7d6" if can_download else "#9aa69c",
        }
        self.resource_download_button.config(**download_button_state)
        if hasattr(self, "software_summary_download_button"):
            self.software_summary_download_button.config(**download_button_state)

    # Показва resource details в интерфейса.
    def _show_resource_details(self) -> None:
        self._refresh_resource_panel()
        messagebox.showinfo(
            "\u0418\u043d\u0441\u0442\u0430\u043b\u0430\u0446\u0438\u043e\u043d\u043d\u0438 \u0440\u0435\u0441\u0443\u0440\u0441\u0438",
            missing_resource_report(self.resource_status),
            parent=self.root,
        )

    # Изтегля missing resources от зададения адрес.
    def _download_missing_resources(self) -> None:
        self._refresh_resource_panel()
        missing_downloads = [
            check.item
            for check in self.resource_status.checks
            if not check.available and check.item.url
        ]
        if not missing_downloads:
            messagebox.showinfo(
                "\u041d\u044f\u043c\u0430 \u0430\u0434\u0440\u0435\u0441\u0438 \u0437\u0430 \u0438\u0437\u0442\u0435\u0433\u043b\u044f\u043d\u0435",
                "Има липсващи ресурси, но в installers_manifest.json още няма зададени URL адреси. "
                "Когато качим пакетите онлайн, добавяме адресите там и бутонът ще започне да ги изтегля.",
                parent=self.root,
            )
            return

        confirmed = messagebox.askyesno(
            "\u0418\u0437\u0442\u0435\u0433\u043b\u044f\u043d\u0435 \u043d\u0430 \u0440\u0435\u0441\u0443\u0440\u0441\u0438",
            f"Да изтегля ли {len(missing_downloads)} липсващи пакета в:\n\n{self.resource_status.installers_root}",
            parent=self.root,
        )
        if not confirmed:
            return

        self.resource_download_button.config(state="disabled")
        if hasattr(self, "software_summary_download_button"):
            self.software_summary_download_button.config(state="disabled")
        self.status_var.set("Изтегляне на липсващи инсталационни ресурси...")
        progress_ui = self._open_resource_download_window(len(missing_downloads))
        threading.Thread(target=self._run_resource_downloads, args=(missing_downloads, progress_ui), daemon=True).start()

    # Стартира resource downloads и връща резултата.
    def _run_resource_downloads(self, items: list[object]) -> None:
        errors: list[str] = []

        # Помощна функция за progress.
        def progress(downloaded: int, total: int, name: str) -> None:
            percent = int((downloaded / total) * 100) if total else 0
            self.root.after(0, lambda: self.status_var.set(f"Изтегляне: {name} - {percent}%"))

        for item in items:
            try:
                download_resource(PROJECT_ROOT, item, progress)
            except Exception as exc:
                errors.append(f"{item.name}: {exc}")

        # Помощна функция за finish.
        def finish() -> None:
            self._refresh_resource_panel()
            if errors:
                messagebox.showerror(
                    "\u041f\u0440\u043e\u0431\u043b\u0435\u043c \u043f\u0440\u0438 \u0438\u0437\u0442\u0435\u0433\u043b\u044f\u043d\u0435",
                    "\n".join(errors),
                    parent=self.root,
                )
                self.status_var.set("Изтеглянето приключи с проблем.")
            else:
                messagebox.showinfo(
                    "\u0413\u043e\u0442\u043e\u0432\u043e",
                    "Липсващите инсталационни ресурси са изтеглени успешно.",
                    parent=self.root,
                )
                self.status_var.set("Инсталационните ресурси са обновени.")
                if self.current_menu == "office_install_center":
                    self._render_cards()

        self.root.after(0, finish)

    # Стартира resource downloads и връща резултата.
    def _run_resource_downloads(self, items: list[object]) -> None:
        errors: list[str] = []
        total_items = len(items)

        progress_window = tk.Toplevel(self.root)
        progress_window.title("Изтегляне на инсталационни ресурси")
        progress_window.geometry("620x320")
        progress_window.configure(bg="#07100a")
        progress_window.resizable(False, False)
        progress_window.transient(self.root)
        apply_app_icon(progress_window)

        wrapper = tk.Frame(progress_window, bg="#07100a", padx=22, pady=18)
        wrapper.pack(fill="both", expand=True)
        tk.Label(wrapper, text="Изтегляне на ресурси", font=("Segoe UI Semibold", 17), fg="#eaffef", bg="#07100a").pack(anchor="w")

        package_var = tk.StringVar(value=f"Подготовка на {total_items} пакета...")
        detail_var = tk.StringVar(value="Очакване на първия пакет.")
        speed_var = tk.StringVar(value="Скорост: - | Оставащо време: -")
        total_var = tk.StringVar(value=f"Пакети: 0/{total_items}")
        for variable, font, color in (
            (package_var, ("Segoe UI Semibold", 11), "#c9ffd0"),
            (detail_var, ("Segoe UI", 10), "#9bc39e"),
            (speed_var, ("Segoe UI", 10), "#ffe08a"),
            (total_var, ("Segoe UI", 10), "#d7f1ff"),
        ):
            tk.Label(wrapper, textvariable=variable, font=font, fg=color, bg="#07100a", anchor="w", justify="left", wraplength=560).pack(anchor="w", fill="x", pady=(8, 0))

        current_progress = tk.IntVar(value=0)
        total_progress = tk.IntVar(value=0)
        ttk.Progressbar(wrapper, maximum=100, variable=current_progress, length=560).pack(fill="x", pady=(14, 6))
        ttk.Progressbar(wrapper, maximum=100, variable=total_progress, length=560).pack(fill="x", pady=(4, 10))

        log_box = tk.Text(wrapper, height=5, bg="#102515", fg="#e7ffe9", insertbackground="#e7ffe9", relief="flat", wrap="word", font=("Consolas", 9))
        log_box.pack(fill="both", expand=True)
        log_box.insert("end", "Стартиране на изтеглянето...\n")
        log_box.config(state="disabled")
        self.root.after(0, lambda: self._center_window(progress_window, 620, 320))

        # Помощна функция за append log.
        def append_log(text: str) -> None:
            if not progress_window.winfo_exists():
                return
            log_box.config(state="normal")
            log_box.insert("end", f"{text}\n")
            log_box.see("end")
            log_box.config(state="disabled")

        # Помощна функция за progress.
        def progress(downloaded: int, total: int, name: str, speed: float = 0.0, eta: int = 0, item_index: int = 1) -> None:
            percent = int((downloaded / total) * 100) if total else 0
            total_percent = int(((item_index - 1) + (percent / 100)) * 100 / max(1, total_items))

            # Помощна функция за update.
            def update() -> None:
                self.status_var.set(f"Изтегляне: {name} - {percent}%")
                if progress_window.winfo_exists():
                    current_progress.set(percent)
                    total_progress.set(total_percent)
                    package_var.set(f"Пакет {item_index}/{total_items}: {name}")
                    detail_var.set(f"{format_file_size(downloaded)} от {format_file_size(total)} ({percent}%)")
                    speed_var.set(f"Скорост: {format_bytes_per_second(speed)} | Оставащо време: {format_duration(eta)}")
                    total_var.set(f"Общ прогрес: {total_percent}% | Пакети: {item_index}/{total_items}")

            self.root.after(0, update)

        for index, item in enumerate(items, start=1):
            try:
                self.root.after(0, lambda item=item, index=index: append_log(f"[{index}/{total_items}] Изтегляне: {item.name}"))
                download_resource(PROJECT_ROOT, item, lambda downloaded, total, name, speed=0.0, eta=0: progress(downloaded, total, name, speed, eta, index))
                self.root.after(0, lambda item=item: append_log(f"Готово: {item.name}"))
            except Exception as exc:
                errors.append(f"{item.name}: {exc}")
                self.root.after(0, lambda item=item, exc=exc: append_log(f"Проблем: {item.name} - {exc}"))

        # Помощна функция за finish.
        def finish() -> None:
            self._refresh_resource_panel()
            if progress_window.winfo_exists():
                total_progress.set(100)
            if errors:
                messagebox.showerror("Проблем при изтегляне", "\n".join(errors), parent=self.root)
                self.status_var.set("Изтеглянето приключи с проблем.")
            else:
                if progress_window.winfo_exists():
                    package_var.set("Всички ресурси са изтеглени успешно.")
                    detail_var.set("Архивите са разархивирани в Installers папката.")
                    speed_var.set("Готово.")
                messagebox.showinfo("Готово", "Липсващите инсталационни ресурси са изтеглени успешно.", parent=self.root)
                self.status_var.set("Инсталационните ресурси са обновени.")
                if self.current_menu == "office_install_center":
                    self._render_cards()

        self.root.after(0, finish)

    # Отваря resource download window или съответния прозорец.
    def _open_resource_download_window(self, total_items: int) -> dict[str, object]:
        # Този прозорец само показва информация. Самото теглене върви отделно.
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Изтегляне на инсталационни ресурси")
        progress_window.geometry("660x350")
        progress_window.configure(bg="#07100a")
        progress_window.resizable(False, False)
        progress_window.transient(self.root)
        apply_app_icon(progress_window)

        wrapper = tk.Frame(progress_window, bg="#07100a", padx=22, pady=18)
        wrapper.pack(fill="both", expand=True)
        tk.Label(
            wrapper,
            text="Изтегляне на ресурси",
            font=("Segoe UI Semibold", 17),
            fg="#eaffef",
            bg="#07100a",
        ).pack(anchor="w")

        package_var = tk.StringVar(value=f"Подготовка на {total_items} пакета...")
        detail_var = tk.StringVar(value="Може да затворите този прозорец. Изтеглянето ще продължи.")
        speed_var = tk.StringVar(value="Скорост: - | Оставащо време: -")
        total_var = tk.StringVar(value=f"Пакети: 0/{total_items}")

        for variable, font, color in (
            (package_var, ("Segoe UI Semibold", 11), "#c9ffd0"),
            (detail_var, ("Segoe UI", 10), "#9bc39e"),
            (speed_var, ("Segoe UI", 10), "#ffe08a"),
            (total_var, ("Segoe UI", 10), "#d7f1ff"),
        ):
            tk.Label(
                wrapper,
                textvariable=variable,
                font=font,
                fg=color,
                bg="#07100a",
                anchor="w",
                justify="left",
                wraplength=600,
            ).pack(anchor="w", fill="x", pady=(8, 0))

        current_progress = tk.IntVar(value=0)
        total_progress = tk.IntVar(value=0)
        ttk.Progressbar(wrapper, maximum=100, variable=current_progress, length=600).pack(fill="x", pady=(14, 6))
        ttk.Progressbar(wrapper, maximum=100, variable=total_progress, length=600).pack(fill="x", pady=(4, 10))

        log_box = tk.Text(
            wrapper,
            height=5,
            bg="#102515",
            fg="#e7ffe9",
            insertbackground="#e7ffe9",
            relief="flat",
            wrap="word",
            font=("Consolas", 9),
        )
        log_box.pack(fill="both", expand=True)
        log_box.insert("end", "Стартиране на изтеглянето...\n")
        log_box.config(state="disabled")

        # Ако прозорецът се затвори, не спираме процеса. Просто скриваме визуалната част.
        progress_window.protocol("WM_DELETE_WINDOW", progress_window.destroy)
        self.root.after(0, lambda: self._center_window(progress_window, 660, 350))

        return {
            "window": progress_window,
            "package_var": package_var,
            "detail_var": detail_var,
            "speed_var": speed_var,
            "total_var": total_var,
            "current_progress": current_progress,
            "total_progress": total_progress,
            "log_box": log_box,
            "total_items": total_items,
        }

    # Помощна функция за resource download window exists.
    def _resource_download_window_exists(self, ui: dict[str, object] | None) -> bool:
        if not ui:
            return False
        window = ui.get("window")
        try:
            return bool(window and window.winfo_exists())
        except tk.TclError:
            return False

    # Помощна функция за append resource download log.
    def _append_resource_download_log(self, ui: dict[str, object] | None, text: str) -> None:
        if not self._resource_download_window_exists(ui):
            return
        log_box = ui.get("log_box")
        try:
            log_box.config(state="normal")
            log_box.insert("end", f"{text}\n")
            log_box.see("end")
            log_box.config(state="disabled")
        except tk.TclError:
            return

    # Обновява resource download ui след промяна в състоянието.
    def _update_resource_download_ui(
        self,
        ui: dict[str, object] | None,
        item_index: int,
        total_items: int,
        name: str,
        downloaded: int,
        total: int,
        speed: float = 0.0,
        eta: int = 0,
        phase: str = "",
    ) -> None:
        # Показваме два прогреса: текущ пакет и общ прогрес за всички пакети.
        percent = int((downloaded / total) * 100) if total else 0
        total_percent = int(((item_index - 1) + (percent / 100)) * 100 / max(1, total_items))
        action = phase or "Изтегляне"
        self.status_var.set(f"{action}: {name} - {percent}% | Общо {total_percent}%")

        if not self._resource_download_window_exists(ui):
            return

        try:
            ui["current_progress"].set(percent)
            ui["total_progress"].set(total_percent)
            ui["package_var"].set(f"Пакет {item_index}/{total_items}: {name}")
            ui["detail_var"].set(f"{action}: {format_file_size(downloaded)} от {format_file_size(total)} ({percent}%)")
            if phase:
                ui["speed_var"].set("Моля изчакайте. Файловете се обработват локално.")
            else:
                ui["speed_var"].set(f"Скорост: {format_bytes_per_second(speed)} | Оставащо време: {format_duration(eta)}")
            ui["total_var"].set(f"Общ прогрес: {total_percent}% | Пакети: {item_index}/{total_items}")
        except tk.TclError:
            return

    # Стартира resource downloads и връща резултата.
    def _run_resource_downloads(self, items: list[object], ui: dict[str, object] | None = None) -> None:
        # Тази функция върви във фонов thread, за да не блокира самото приложение.
        errors: list[str] = []
        total_items = len(items)

        for index, item in enumerate(items, start=1):
            self.root.after(0, lambda item=item, index=index: self._append_resource_download_log(ui, f"[{index}/{total_items}] Изтегляне: {item.name}"))

            # Помощна функция за progress.
            def progress(
                downloaded: int,
                total: int,
                name: str,
                speed: float = 0.0,
                eta: int = 0,
                phase: str = "",
                item_index: int = index,
            ) -> None:
                # Всички промени по Tkinter ги връщаме към главния прозорец с root.after.
                self.root.after(
                    0,
                    lambda: self._update_resource_download_ui(
                        ui,
                        item_index,
                        total_items,
                        name,
                        downloaded,
                        total,
                        speed,
                        eta,
                        phase,
                    ),
                )

            try:
                download_resource(PROJECT_ROOT, item, progress)
                self.root.after(0, lambda item=item: self._append_resource_download_log(ui, f"Готово: {item.name}"))
            except Exception as exc:
                errors.append(f"{item.name}: {exc}")
                self.root.after(0, lambda item=item, exc=exc: self._append_resource_download_log(ui, f"Проблем: {item.name} - {exc}"))

        # Помощна функция за finish.
        def finish() -> None:
            self._refresh_resource_panel()
            if self._resource_download_window_exists(ui):
                try:
                    ui["total_progress"].set(100)
                except tk.TclError:
                    pass

            if errors:
                self.status_var.set("Изтеглянето приключи, но има проблеми. Проверете съобщението.")
                messagebox.showerror("Проблем при изтегляне", "\n".join(errors), parent=self.root)
                return

            if self._resource_download_window_exists(ui):
                try:
                    ui["package_var"].set("Всички ресурси са изтеглени успешно.")
                    ui["detail_var"].set("Архивите са разархивирани и готови за използване.")
                    ui["speed_var"].set("Готово.")
                except tk.TclError:
                    pass
            self.status_var.set("Инсталационните ресурси са обновени.")
            if self.current_menu == "office_install_center":
                self._render_cards()
            messagebox.showinfo("Готово", "Липсващите инсталационни ресурси са изтеглени успешно.", parent=self.root)

        self.root.after(0, finish)

    # Зарежда system health async от файл или конфигурация.
    def _load_system_health_async(self) -> None:
        if self.health_refresh_in_progress:
            return
        self.health_refresh_in_progress = True
        threading.Thread(target=self._collect_and_render_system_health, daemon=True).start()

    # Зарежда language status async от файл или конфигурация.
    def _load_language_status_async(self) -> None:
        threading.Thread(target=self._collect_and_render_language_status, daemon=True).start()

    # Събира and render language status от системата.
    def _collect_and_render_language_status(self) -> None:
        try:
            status = get_language_status()
            self.language_status_cache = status
            text = self._build_language_status_summary(status)
            color = "#9aff9f" if status.has_language_pack or status.has_bulgarian else "#ffb0a8"
        except Exception as exc:
            text = f"Езиков статус: грешка при проверка\n{exc}"
            color = "#ffb0a8"
        self.root.after(0, lambda: self._apply_language_status_summary(text, color))

    # Подготвя language status summary според избраните настройки.
    def _build_language_status_summary(self, status: LanguageStatus) -> str:
        # Помощна функция за marker.
        def marker(value: bool) -> str:
            return "[OK]" if value else "[--]"

        return (
            f"{marker(status.has_bulgarian)} bg-BG в списъка\n"
            f"{marker(status.has_language_pack)} Български езиков пакет\n"
            f"{marker(status.has_bds)} БДС клавиатура\n"
            f"{marker(status.has_phonetic)} Фонетична стандартна\n"
            f"{marker(status.has_traditional)} Фонетична традиционна"
        )

    # Помощна функция за apply language status summary.
    def _apply_language_status_summary(self, text: str, color: str) -> None:
        self.language_status_var.set(text)
        self.language_status_label.config(fg=color)

    # Събира and render system health от системата.
    def _collect_and_render_system_health(self) -> None:
        try:
            items = collect_health_items()
        except Exception as exc:
            items = [HealthItem(label="Health:", value=f"Diagnostics failed: {exc}", ok=False)]
        self.latest_health_items = items
        self.root.after(0, lambda: self._apply_system_health_update(items))

    # Помощна функция за apply system health update.
    def _apply_system_health_update(self, items: list[HealthItem]) -> None:
        # Прилага новите health данни, обновява dashboard-а и планира следващ refresh.
        self.health_refresh_in_progress = False
        self._render_system_health(items)
        if self.current_menu == "main":
            self._update_dashboard_live_widgets()
        if self.health_refresh_job is not None:
            self.root.after_cancel(self.health_refresh_job)
        self.health_refresh_job = self.root.after(self.health_refresh_interval_ms, self._load_system_health_async)

    # Рисува system health върху текущия екран.
    def _render_system_health(self, items: list[HealthItem]) -> None:
        if self.health_scroll_job is not None:
            self.root.after_cancel(self.health_scroll_job)
            self.health_scroll_job = None

        for widget in self.health_frame.winfo_children():
            widget.destroy()
        self.health_rows.clear()

        self.health_canvas = tk.Canvas(
            self.health_frame,
            bg="#112716",
            highlightthickness=0,
            bd=0,
        )
        self.health_scrollbar = ttk.Scrollbar(
            self.health_frame,
            orient="vertical",
            command=self.health_canvas.yview,
        )
        self.health_inner_frame = tk.Frame(self.health_canvas, bg="#112716")
        self.health_canvas_window = self.health_canvas.create_window(
            (0, 0),
            window=self.health_inner_frame,
            anchor="nw",
        )
        self.health_canvas.configure(yscrollcommand=self.health_scrollbar.set)
        self.health_canvas.pack(side="left", fill="both", expand=True)
        self.health_scrollbar.pack(side="right", fill="y")

        # Обновява scroll region след промяна в състоянието.
        def update_scroll_region(_: object | None = None) -> None:
            if self.health_canvas is None or self.health_inner_frame is None:
                return
            self.health_canvas.configure(scrollregion=self.health_canvas.bbox("all"))
            self.health_canvas.itemconfigure(self.health_canvas_window, width=self.health_canvas.winfo_width())

        self.health_inner_frame.bind("<Configure>", update_scroll_region)
        self.health_canvas.bind("<Configure>", update_scroll_region)
        self.health_canvas.bind("<MouseWheel>", self._on_health_mousewheel)

        for item in items:
            row = tk.Frame(self.health_inner_frame, bg="#112716")
            row.pack(fill="x", padx=10, pady=4)

            status_text = "OK" if item.ok else "!"
            status_color = "#7dff92" if item.ok else "#ff6f6f"
            status_label = tk.Label(
                row,
                text=status_text,
                font=("Segoe UI Semibold", 11),
                fg=status_color,
                bg="#112716",
                width=2,
                anchor="w",
            )
            status_label.pack(side="left")

            name_label = tk.Label(
                row,
                text=item.label,
                font=("Segoe UI Semibold", 9),
                fg="#d8ffe0",
                bg="#112716",
                width=13,
                anchor="w",
            )
            name_label.pack(side="left")

            value_label = tk.Label(
                row,
                text=item.value,
                font=("Segoe UI", 9),
                fg=status_color,
                bg="#112716",
                justify="left",
                wraplength=160,
                anchor="w",
            )
            value_label.pack(side="left", fill="x", expand=True)
            self.health_rows.append((status_label, name_label, value_label))

        self.health_frame.after(200, self._start_health_auto_scroll)

    # Обработва събитието on health mousewheel.
    def _on_health_mousewheel(self, event: tk.Event) -> None:
        if self.health_canvas is None:
            return
        self.health_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # Обработва събитието on dashboard info mousewheel.
    def _on_dashboard_info_mousewheel(self, canvas: tk.Canvas, event: tk.Event) -> str:
        # Позволява ръчно скролване с мишката в картата със системното състояние.
        self._stop_dashboard_info_scroll()
        info_map = self.dashboard_live_widgets.get("system_info_scroll")
        scroll_args = self.dashboard_live_widgets.get("system_info_scroll_scroll_args", ())
        if not isinstance(info_map, dict):
            return "break"
        if not (isinstance(scroll_args, tuple) and len(scroll_args) == 3):
            return "break"
        delta = getattr(event, "delta", 0)
        if getattr(event, "num", None) == 4:
            delta = 120
        elif getattr(event, "num", None) == 5:
            delta = -120
        if delta:
            item_a, item_b, content_height = scroll_args
            step = max(12, self._scale_px(18))
            offset = step if delta > 0 else -step
            for item_id in (item_a, item_b):
                x_pos, y_pos = canvas.coords(item_id)
                canvas.coords(item_id, x_pos, y_pos + offset)
            a_y = canvas.coords(item_a)[1]
            b_y = canvas.coords(item_b)[1]
            if a_y > content_height:
                canvas.coords(item_a, 0, b_y - content_height)
            if b_y > content_height:
                canvas.coords(item_b, 0, a_y - content_height)
            if a_y <= -content_height:
                canvas.coords(item_a, 0, b_y + content_height)
            if b_y <= -content_height:
                canvas.coords(item_b, 0, a_y + content_height)
        self.dashboard_info_scroll_job = self.root.after(
            1800,
            lambda: self._start_dashboard_info_scroll(canvas, *self.dashboard_live_widgets.get("system_info_scroll_scroll_args", ())),
        )
        return "break"

    # Помощна функция за start health auto scroll.
    def _start_health_auto_scroll(self) -> None:
        if self.health_canvas is None:
            return
        self.health_canvas.update_idletasks()
        first, last = self.health_canvas.yview()
        if first <= 0.0 and last >= 1.0:
            return
        self.health_scroll_position = 0.0
        self.health_canvas.yview_moveto(0.0)
        self.health_scroll_job = self.root.after(1200, self._auto_scroll_health)

    # Помощна функция за auto scroll health.
    def _auto_scroll_health(self) -> None:
        if self.health_canvas is None or not self.health_canvas.winfo_exists():
            self.health_scroll_job = None
            return

        _, last = self.health_canvas.yview()
        if last >= 0.995:
            self.health_scroll_position = 0.0
            self.health_canvas.yview_moveto(0.0)
            delay = 1700
        else:
            self.health_scroll_position = min(1.0, self.health_scroll_position + 0.003)
            self.health_canvas.yview_moveto(self.health_scroll_position)
            delay = 80

        self.health_scroll_job = self.root.after(delay, self._auto_scroll_health)

    # Помощна функция за stop dashboard info scroll.
    def _stop_dashboard_info_scroll(self) -> None:
        # Спира скрола на картата със системната информация.
        if self.dashboard_info_scroll_job is not None:
            try:
                self.root.after_cancel(self.dashboard_info_scroll_job)
            except tk.TclError:
                pass
            self.dashboard_info_scroll_job = None
        self.dashboard_info_scroll_position = 0.0

    # Обновява dashboard live widgets след промяна в състоянието.
    def _update_dashboard_live_widgets(self) -> None:
        # Обновява само живите карти в dashboard-а, без да прерисува цялото меню.
        if self.current_menu != "main" or not self.dashboard_live_widgets:
            return

        metric_specs = {
            str(spec["key"]): spec
            for spec in self._dashboard_metric_cards()
        }
        for key, spec in metric_specs.items():
            widget_map = self.dashboard_live_widgets.get(key)
            if not isinstance(widget_map, dict):
                continue
            value_label = widget_map.get("value_label")
            status_label = widget_map.get("status_label")
            percent_label = widget_map.get("percent_label")
            fill_widget = widget_map.get("fill_widget")
            track_widget = widget_map.get("track_widget")
            ok = bool(spec["ok"])
            accent = APP_ACCENT if ok else APP_WARNING
            value_fg = APP_ACCENT if ok else "#ff8b8b"
            if isinstance(value_label, tk.Label) and value_label.winfo_exists():
                value_label.config(text=str(spec["value"]), fg=value_fg)
            if isinstance(status_label, tk.Label) and status_label.winfo_exists():
                status_label.config(text=str(spec["status"]))
            percent_value = self._dashboard_metric_percent(str(spec["value"]), ok)
            if isinstance(percent_label, tk.Label) and percent_label.winfo_exists():
                percent_label.config(text=f"{percent_value}%", fg=accent)
            if isinstance(fill_widget, tk.Frame) and fill_widget.winfo_exists():
                fill_widget.configure(bg=accent)
                track_width = track_widget.winfo_width() if isinstance(track_widget, tk.Frame) and track_widget.winfo_exists() else 220
                fill_width = max(10, int(track_width * (percent_value / 100)))
                fill_widget.place_configure(width=fill_width)

        alert_widgets = self.dashboard_live_widgets.get("security_alert")
        if isinstance(alert_widgets, dict):
            problem_count = sum(1 for item in self.latest_health_items if not item.ok)
            alert_ok = problem_count == 0
            title_label = alert_widgets.get("title_label")
            detail_label = alert_widgets.get("detail_label")
            strip_widget = alert_widgets.get("strip_widget")
            accent = APP_ACCENT if alert_ok else APP_DANGER
            if isinstance(strip_widget, tk.Frame) and strip_widget.winfo_exists():
                strip_widget.configure(bg=accent)
            if isinstance(title_label, tk.Label) and title_label.winfo_exists():
                title_label.config(text="Няма проблем" if alert_ok else "Внимание", fg=accent)
            if isinstance(detail_label, tk.Label) and detail_label.winfo_exists():
                detail_label.config(
                    text="Системата изглежда стабилна" if alert_ok else f"{problem_count} проблем(а) открити",
                    fg=APP_TEXT_SOFT if alert_ok else "#ffd4d4",
                )

        info_widgets = self.dashboard_live_widgets.get("system_info_scroll")
        if isinstance(info_widgets, dict):
            self._refresh_dashboard_info_panel(info_widgets)

    # Помощна функция за refresh dashboard info panel.
    def _refresh_dashboard_info_panel(self, widget_map: dict[str, object]) -> None:
        # Обновява само картата със системното състояние.
        canvas = widget_map.get("canvas")
        frame_a = widget_map.get("frame_a")
        frame_b = widget_map.get("frame_b")
        refresh_callback = widget_map.get("refresh_callback")
        if not isinstance(canvas, tk.Canvas) or not canvas.winfo_exists():
            return
        if not isinstance(frame_a, tk.Frame) or not frame_a.winfo_exists():
            return
        if not isinstance(frame_b, tk.Frame) or not frame_b.winfo_exists():
            return

        self._stop_dashboard_info_scroll()
        for child in frame_a.winfo_children():
            child.destroy()
        for child in frame_b.winfo_children():
            child.destroy()
        self._populate_dashboard_info_rows(frame_a, 235)
        self._populate_dashboard_info_rows(frame_b, 235)
        self._bind_dashboard_info_mousewheel(frame_a, canvas)
        self._bind_dashboard_info_mousewheel(frame_b, canvas)
        canvas.update_idletasks()
        if callable(refresh_callback):
            refresh_callback()

    # Помощна функция за bind dashboard info mousewheel.
    def _bind_dashboard_info_mousewheel(self, widget: tk.Misc, canvas: tk.Canvas) -> None:
        # Връзва колелцето към всички вътрешни елементи на картата.
        for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            widget.bind(sequence, lambda event, target=canvas: self._on_dashboard_info_mousewheel(target, event))
        for child in widget.winfo_children():
            self._bind_dashboard_info_mousewheel(child, canvas)

    # Помощна функция за populate dashboard info rows.
    def _populate_dashboard_info_rows(self, parent: tk.Frame, wraplength: int) -> None:
        # Пълни плаващата карта с редовете за системата.
        system_icon_map = {
            "Общо състояние": "shield_small",
            "Компютър": "monitor_small",
            "Потребител": "admin_small",
            "Компютър / потребител": "admin_small",
            "Операционна система": "shield_small",
            "Процесор": "cpu_small",
            "Натоварване на процесора": "cpu_small",
            "Температура на процесора": "cpu_small",
            "Напрежение на процесора": "bolt_small",
            "Дънна платка": "drive_small",
            "RAM използване": "ram_small",
            "RAM тип и скорост": "ram_small",
            "Графична карта": "bolt_small",
            "BIOS версия": "shield_small",
            "Време на работа": "refresh_small",
            "IP адрес": "network_small",
            "Secure Boot": "shield_small",
            "Батерия": "shield_small",
            "Дискове": "drive_small",
        }
        for label, value in self._dashboard_system_rows():
            row = tk.Frame(parent, bg=APP_PANEL_ALT)
            row.pack(fill="x", padx=12, pady=4)
            inner = tk.Frame(row, bg=APP_PANEL_ALT)
            inner.pack(fill="x", padx=10, pady=7)
            row_icon = self.dashboard_icons.get(system_icon_map.get(label, "shield_small"))
            if row_icon is not None:
                tk.Label(inner, image=row_icon, bg=APP_PANEL_ALT).pack(side="left", padx=(0, 8))
            else:
                tk.Label(inner, text="•", font=self._font(8, "bold", "Segoe UI Semibold"), fg=APP_ACCENT, bg=APP_PANEL_ALT).pack(side="left", padx=(0, 8))
            value_color = APP_ACCENT if label == "Общо състояние" and "OK" in value else (APP_DANGER if label == "Общо състояние" else APP_TEXT)
            tk.Label(
                inner,
                text=f"{label}:",
                font=self._font(7),
                fg=APP_TEXT_MUTED,
                bg=APP_PANEL_ALT,
                width=16,
                anchor="w",
            ).pack(side="left")
            tk.Label(
                inner,
                text=value,
                font=self._font(7, "bold" if label == "Общо състояние" else "normal", "Segoe UI Semibold" if label == "Общо състояние" else "Segoe UI"),
                fg=value_color,
                bg=APP_PANEL_ALT,
                justify="left",
                anchor="w",
                wraplength=wraplength,
            ).pack(side="left", fill="x", expand=True)

    # Помощна функция за start dashboard info scroll.
    def _start_dashboard_info_scroll(self, canvas: tk.Canvas, item_a: int, item_b: int, content_height: int) -> None:
        # Движи текста нагоре плавно в картата със системното състояние.
        if self.current_menu != "main" or not canvas.winfo_exists() or content_height <= 0:
            self.dashboard_info_scroll_job = None
            return

        visible_height = max(canvas.winfo_height(), 1)
        if content_height <= visible_height:
            self.dashboard_info_scroll_job = None
            return

        step = max(1, self._scale_px(1))
        for item_id in (item_a, item_b):
            x_pos, y_pos = canvas.coords(item_id)
            canvas.coords(item_id, x_pos, y_pos - step)

        a_y = canvas.coords(item_a)[1]
        b_y = canvas.coords(item_b)[1]
        if a_y <= -content_height:
            canvas.coords(item_a, 0, b_y + content_height)
        elif b_y <= -content_height:
            canvas.coords(item_b, 0, a_y + content_height)

        self.dashboard_info_scroll_job = self.root.after(
            38,
            lambda: self._start_dashboard_info_scroll(canvas, item_a, item_b, content_height),
        )

    # Обработва събитието on dashboard info mousewheel.
    def _on_dashboard_info_mousewheel(self, canvas: tk.Canvas, event: tk.Event) -> str:
        # Позволява скрол с мишката в картата със системното състояние.
        self._stop_dashboard_info_scroll()
        if not canvas.winfo_exists():
            return "break"
        delta = getattr(event, "delta", 0)
        if getattr(event, "num", None) == 4:
            delta = 120
        elif getattr(event, "num", None) == 5:
            delta = -120
        if delta:
            direction = -1 if delta > 0 else 1
            canvas.yview_scroll(direction, "units")
        self.dashboard_info_scroll_job = self.root.after(
            1800,
            lambda: self._start_dashboard_info_scroll(canvas),
        )
        return "break"

    # Помощна функция за refresh dashboard info panel.
    def _refresh_dashboard_info_panel(self, widget_map: dict[str, object]) -> None:
        # Обновява само картата със системното състояние.
        canvas = widget_map.get("canvas")
        frame = widget_map.get("frame")
        refresh_callback = widget_map.get("refresh_callback")
        if not isinstance(canvas, tk.Canvas) or not canvas.winfo_exists():
            return
        if not isinstance(frame, tk.Frame) or not frame.winfo_exists():
            return
        self._stop_dashboard_info_scroll()
        for child in frame.winfo_children():
            child.destroy()
        self._populate_dashboard_info_rows(frame, 235)
        self._bind_dashboard_info_mousewheel(frame, canvas)
        canvas.update_idletasks()
        if callable(refresh_callback):
            refresh_callback()

    # Помощна функция за start dashboard info scroll.
    def _start_dashboard_info_scroll(self, canvas: tk.Canvas) -> None:
        # Движи текста нагоре плавно в картата със системното състояние.
        if self.current_menu != "main" or not canvas.winfo_exists():
            self.dashboard_info_scroll_job = None
            return
        canvas.update_idletasks()
        first, last = canvas.yview()
        if first <= 0.0 and last >= 1.0:
            self.dashboard_info_scroll_job = None
            return
        if last >= 0.995:
            canvas.yview_moveto(0.0)
            delay = 1400
        else:
            canvas.yview_scroll(1, "units")
            delay = 85
        self.dashboard_info_scroll_job = self.root.after(
            delay,
            lambda: self._start_dashboard_info_scroll(canvas),
        )

    # Обработва събитието on dashboard canvas mousewheel.
    def _on_dashboard_canvas_mousewheel(self, canvas: tk.Canvas, event: tk.Event) -> str:
        # Позволява обикновен скрол с мишката в dashboard canvas без автоматично движение.
        if not canvas.winfo_exists():
            return "break"
        delta = getattr(event, "delta", 0)
        if getattr(event, "num", None) == 4:
            delta = 120
        elif getattr(event, "num", None) == 5:
            delta = -120
        if delta:
            direction = -1 if delta > 0 else 1
            canvas.yview_scroll(direction, "units")
        return "break"

    # Помощна функция за bind dashboard canvas mousewheel.
    def _bind_dashboard_canvas_mousewheel(self, widget: tk.Misc, canvas: tk.Canvas) -> None:
        # Връзва колелцето към всички вътрешни елементи на scrollable canvas в dashboard-а.
        for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            widget.bind(sequence, lambda event, target=canvas: self._on_dashboard_canvas_mousewheel(target, event))
        for child in widget.winfo_children():
            self._bind_dashboard_canvas_mousewheel(child, canvas)

    # Проверява updates async и връща резултат за интерфейса.
    def _check_updates_async(self) -> None:
        threading.Thread(target=self._perform_update_check, daemon=True).start()

    # Помощна функция за perform update check.
    def _perform_update_check(self) -> None:
        result = check_for_updates(
            str(self.version_info["version"]),
            str(self.version_info.get("version_info_url", "")),
        )
        self.root.after(0, lambda: self._apply_update_result(result))

    # Помощна функция за apply update result.
    def _apply_update_result(self, result: UpdateResult) -> None:
        self.update_result = result
        status_map = {
            "checking": {
                "icon": "i",
                "bg": "#153042",
                "border": "#2a5975",
                "fg": "#d7f1ff",
                "button_bg": "#2b607a",
                "button_text": "",
                "button_state": "disabled",
                "message": f"Проверка за актуализации за v{self.version_info['version']}...",
            },
            "up_to_date": {
                "icon": "\u2713",
                "bg": "#14301d",
                "border": "#2f6a40",
                "fg": "#d8ffe3",
                "button_bg": "#245634",
                "button_text": "",
                "button_state": "disabled",
                "message": f"Приложението е актуално. Текуща версия: v{self.version_info['version']}.",
            },
            "update_available": {
                "icon": "\u2191",
                "bg": "#423014",
                "border": "#8a6a2a",
                "fg": "#ffeec5",
                "button_bg": "#8a6a2a",
                "button_text": "Инсталирай",
                "button_state": "normal",
                "message": f"Налична е нова версия: v{result.latest_version}. {result.notes or 'Има по-нова версия в GitHub.'}",
            },
            "not_configured": {
                "icon": "!",
                "bg": "#352b13",
                "border": "#7d6a2d",
                "fg": "#f9e6a8",
                "button_bg": "#7d6a2d",
                "button_text": "",
                "button_state": "disabled",
                "message": "Онлайн проверката не е конфигурирана. Добави GitHub raw адрес към version.json.",
            },
            "raw_unavailable": {
                "icon": "!",
                "bg": "#352b13",
                "border": "#7d6a2d",
                "fg": "#f9e6a8",
                "button_bg": "#7d6a2d",
                "button_text": "",
                "button_state": "disabled",
                "message": "Онлайн проверката е конфигурирана, но GitHub raw version.json не е публично достъпен. Провери дали repo-то е Public и дали файлът version.json е качен в main.",
            },
            "error": {
                "icon": "\u2717",
                "bg": "#411717",
                "border": "#7b2d2d",
                "fg": "#ffc8c8",
                "button_bg": "#7b2d2d",
                "button_text": "",
                "button_state": "disabled",
                "message": f"Проверката за актуализация не успя: {result.error or 'неизвестна грешка'}",
            },
        }
        style = status_map.get(result.status, status_map["error"])
        self.update_download_url = result.download_url or self.version_info.get("download_url", "")
        self.update_package_url = result.package_url

        self.update_banner.config(bg=style["bg"], highlightbackground=style["border"])
        self.update_icon_label.config(text=style["icon"], bg=style["bg"], fg=style["fg"])
        self.update_message_var.set(style["message"])
        self.update_message_label.config(bg=style["bg"], fg=style["fg"])
        self.update_action_button.config(
            text=style["button_text"] or "Отвори",
            bg=style["button_bg"],
            activebackground=style["button_bg"],
            state=style["button_state"],
        )

        self.update_history_button.config(bg=style["button_bg"], activebackground=style["button_bg"])

        if result.status == "update_available" and not self.update_popup_shown:
            self.update_popup_shown = True
            self.root.after(250, lambda: self._show_update_available_dialog(result))

    # Обновява history lines след промяна в състоянието.
    def _update_history_lines(self) -> list[str]:
        if self.update_result and self.update_result.changelog:
            return list(self.update_result.changelog)
        raw_changelog = self.version_info.get("changelog", [])
        if isinstance(raw_changelog, list):
            return [str(item) for item in raw_changelog if str(item).strip()]
        return []

    # Показва update available dialog в интерфейса.
    def _show_update_available_dialog(self, result: UpdateResult) -> None:
        details = "\n".join(f"- {item}" for item in (result.changelog or ())[:6])
        message = (
            f"Available update: v{result.latest_version}\n\n"
            f"{result.notes or 'A newer version is available on GitHub.'}"
        )
        if details:
            message += f"\n\nChangelog:\n{details}"
        if messagebox.askyesno("Available Update", f"{message}\n\nInstall it now?", parent=self.root):
            self._open_update_download()

    # Показва update history в интерфейса.
    def _show_update_history(self) -> None:
        history_window = tk.Toplevel(self.root)
        history_window.title("Update History")
        history_window.geometry("620x430")
        history_window.transient(self.root)
        apply_app_icon(history_window)

        wrapper = tk.Frame(history_window, bg="#0d1711", padx=18, pady=16)
        wrapper.pack(fill="both", expand=True)

        tk.Label(
            wrapper,
            text="История на актуализациите",
            font=("Segoe UI Semibold", 15),
            bg="#0d1711",
            fg="#effff2",
        ).pack(anchor="w")

        latest = self.update_result.latest_version if self.update_result else str(self.version_info.get("version", ""))
        status = "Няма намерена нова версия."
        if self.update_result and self.update_result.status == "update_available":
            status = f"Налична е нова версия: v{latest}"
        elif self.update_result and self.update_result.status == "up_to_date":
            status = f"Приложението е актуално: v{self.version_info.get('version', '')}"

        tk.Label(
            wrapper,
            text=status,
            font=("Segoe UI", 10),
            bg="#0d1711",
            fg="#aee8b8",
        ).pack(anchor="w", pady=(4, 12))

        text_box = tk.Text(
            wrapper,
            bg="#07100a",
            fg="#e7ffe9",
            insertbackground="#e7ffe9",
            relief="flat",
            wrap="word",
            font=("Segoe UI", 10),
            padx=12,
            pady=12,
        )
        text_box.pack(fill="both", expand=True)

        lines = self._update_history_lines()
        if lines:
            text_box.insert("end", "\n\n".join(lines))
        else:
            text_box.insert("end", "Все още няма добавена история на актуализациите.")
        text_box.config(state="disabled")

        bottom = tk.Frame(wrapper, bg="#0d1711")
        bottom.pack(fill="x", pady=(12, 0))
        tk.Button(
            bottom,
            text="Затвори",
            command=history_window.destroy,
            font=("Segoe UI Semibold", 10),
            bg="#245634",
            fg="#f3fff5",
            activebackground="#2f7044",
            activeforeground="#ffffff",
            bd=0,
            padx=18,
            pady=8,
            cursor="hand2",
        ).pack(side="right")

    # Отваря update download или съответния прозорец.
    def _open_update_download(self) -> None:
        package_url = self.update_package_url.strip()
        if package_url:
            self._install_update_package(package_url)
            return

        if not self.update_download_url.strip():
            messagebox.showinfo(
                "Няма адрес за изтегляне",
                "За тази актуализация не е зададен автоматичен update пакет.",
                parent=self.root,
            )
            return
        webbrowser.open(self.update_download_url.strip())

    # Помощна функция за restart command.
    def _restart_command(self) -> list[str]:
        if getattr(sys, "frozen", False):
            return [sys.executable]
        return [sys.executable, str(Path(__file__).resolve())]

    # Стартира инсталационната логика за update package.
    def _install_update_package(self, package_url: str) -> None:
        if self.update_installing:
            return
        if not messagebox.askyesno(
            "Install Update",
            "The application will download the update, replace the files, close, and reopen. Continue?",
            parent=self.root,
        ):
            return

        self.update_installing = True
        progress_window = tk.Toplevel(self.root)
        progress_window.title("WGA Update")
        progress_window.transient(self.root)
        progress_window.resizable(False, False)
        apply_app_icon(progress_window)

        panel = tk.Frame(progress_window, bg="#101820", padx=22, pady=18)
        panel.pack(fill="both", expand=True)
        tk.Label(
            panel,
            text="Инсталиране на актуализация",
            font=("Segoe UI Semibold", 13),
            bg="#101820",
            fg="#f1fff5",
        ).pack(anchor="w")
        status_var = tk.StringVar(value="Подготовка...")
        tk.Label(
            panel,
            textvariable=status_var,
            font=("Segoe UI", 10),
            bg="#101820",
            fg="#b9d8c3",
            wraplength=420,
            justify="left",
        ).pack(anchor="w", pady=(8, 12))
        progress_var = tk.IntVar(value=0)
        progress_bar = ttk.Progressbar(panel, maximum=100, variable=progress_var, length=420)
        progress_bar.pack(fill="x")

        progress_window.update_idletasks()
        width = progress_window.winfo_width()
        height = progress_window.winfo_height()
        x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - width) // 2)
        y = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - height) // 2)
        progress_window.geometry(f"+{x}+{y}")

        # Обновява progress след промяна в състоянието.
        def update_progress(value: int) -> None:
            self.root.after(0, lambda: progress_var.set(value))

        # Помощна функция за worker.
        def worker() -> None:
            try:
                self.root.after(0, lambda: status_var.set("Сваляне на update пакета..."))
                helper_path = prepare_update_install(
                    package_url=package_url,
                    target_root=PROJECT_ROOT,
                    restart_command=self._restart_command(),
                    progress_callback=update_progress,
                )

                # Помощна функция за finish.
                def finish() -> None:
                    status_var.set("Готово. Приложението ще се рестартира...")
                    progress_var.set(100)
                    launch_helper_and_exit(helper_path)
                    self.root.after(500, self.root.destroy)

                self.root.after(0, finish)
            except Exception as exc:
                # Помощна функция за fail.
                def fail() -> None:
                    self.update_installing = False
                    progress_window.destroy()
                    messagebox.showerror(
                        "Грешка при актуализация",
                        f"Актуализацията не успя:\n{exc}",
                        parent=self.root,
                    )

                self.root.after(0, fail)

        threading.Thread(target=worker, daemon=True).start()

    # Помощна функция за make nav button.
    def _make_nav_button(
        self,
        parent: tk.Widget,
        text: str,
        command: object,
        accent: str = APP_PANEL_ALT,
    ) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            font=self._font(10, "bold", "Segoe UI Semibold"),
            bg=accent,
            fg="#effff8",
            activebackground=APP_BORDER_STRONG,
            activeforeground="#ffffff",
            bd=0,
            highlightthickness=1,
            highlightbackground=APP_BORDER,
            padx=18,
            pady=10,
            width=self.nav_button_char_width,
            cursor="hand2",
        )

    # Помощна функция за make card button.
    def _make_card_button(
        self,
        parent: tk.Widget,
        text: str,
        command: object,
        bg: str,
        fg: str,
        active_bg: str,
        *,
        state: str = "normal",
        cursor: str = "hand2",
    ) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            font=self._font(10, "bold", "Segoe UI Semibold"),
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground="#ffffff",
            bd=0,
            borderwidth=0,
            relief="flat",
            overrelief="flat",
            highlightthickness=1,
            highlightbackground=APP_BORDER,
            takefocus=0,
            padx=16,
            pady=8,
            width=CARD_BUTTON_WIDTH,
            height=CARD_BUTTON_HEIGHT,
            wraplength=max(200, self.card_button_width_px - 28),
            justify="center",
            cursor=cursor,
            state=state,
        )

    # Рисува menu върху текущия екран.
    def render_menu(self, menu_key: str, reset_history: bool = False) -> None:
        if reset_history:
            self.history.clear()
        self.current_menu = menu_key
        self.current_page = 0
        menu = MENU_TREE[menu_key]

        self.menu_path.config(text=self._build_path())
        self.card_title.config(text=menu["title"])
        self.card_subtitle.config(text=menu["subtitle"])
        self.subtitle_label.config(text=menu["subtitle"])
        self.status_var.set(f"Отворено е меню: {menu['title']}.")
        self.header_dashboard_button.config(state="normal")
        self.dashboard_button.config(state="disabled" if menu_key == "main" else "normal")
        self._refresh_sidebar_navigation()
        self._refresh_overview_cards()
        install_style_menus = {"install_software", "office_install_center", "secret_install", "office_center"}
        compact_content_menus = install_style_menus | {"auto_installer", "driver_backup", "language", "nexus_admin"}
        use_install_style = menu_key in install_style_menus
        use_compact_content = menu_key in compact_content_menus
        self._toggle_dashboard_chrome(
            menu_key == "main",
            hide_overview=self._is_activation_menu(menu_key) or use_compact_content,
            hide_update_banner=self._is_activation_menu(menu_key) or use_compact_content,
            show_resource_panel=False,
            show_software_summary=use_install_style,
        )
        self._toggle_language_status_panel(False)
        self._render_cards()
        if menu_key == "main":
            self.root.after(120, self._ensure_dashboard_visible)

    # Помощна функция за toggle language status panel.
    def _toggle_language_status_panel(self, visible: bool) -> None:
        if not visible and self.language_status_panel.winfo_ismapped():
            self.language_status_panel.grid_forget()

    # Подготвя path според избраните настройки.
    def _build_path(self) -> str:
        trail = [MENU_TREE[key]["title"] for key in self.history + [self.current_menu]]
        return " > ".join(trail)

    # Рисува cards върху текущия екран.
    def _render_cards(self) -> None:
        if self.current_menu == "main" and self.dashboard_is_rendering:
            self.root.after(50, self._render_cards)
            return
        self._stop_dashboard_info_scroll()
        if self.dashboard_render_job is not None:
            try:
                self.root.after_cancel(self.dashboard_render_job)
            except tk.TclError:
                pass
            self.dashboard_render_job = None
        self.dashboard_live_widgets = {}
        for widget in self.cards_frame.winfo_children():
            if widget is self.language_status_panel:
                widget.grid_forget()
                continue
            widget.destroy()
        self.dashboard_host_frame = None
        for index in range(12):
            self.cards_frame.rowconfigure(index, weight=0, minsize=0)
            self.cards_frame.columnconfigure(index, weight=0, minsize=0)

        if self.current_menu == "auto_installer":
            self._render_auto_installer()
            return
        if self.current_menu == "main":
            self._render_home_dashboard()
            return
        if self.current_menu == "activation":
            self._render_activation_menu()
            return

        items = MENU_TREE[self.current_menu]["items"]
        page_size = MENU_PAGE_SIZE.get(self.current_menu, CARDS_PER_PAGE)
        total_pages = max(1, math.ceil(len(items) / page_size))
        self.current_page = max(0, min(self.current_page, total_pages - 1))
        start = self.current_page * page_size
        page_items = items[start : start + page_size]

        card_columns = self._card_columns()
        for column in range(card_columns):
            self.cards_frame.columnconfigure(column, weight=1, uniform="cards")
        row_count = max(1, math.ceil(len(page_items) / card_columns))
        row_minsize = self.scaled_menu_card_min_height.get(self.current_menu, self.scaled_card_min_height)
        for row in range(row_count):
            self.cards_frame.rowconfigure(row, weight=1, minsize=row_minsize)

        for index, item in enumerate(page_items):
            row = index // card_columns
            column = index % card_columns
            card = self._build_card(self.cards_frame, item)
            card.grid(row=row, column=column, sticky="nsew", padx=8, pady=8)

        if self.language_status_panel.winfo_ismapped():
            self.language_status_panel.grid_forget()

        self.page_label.config(text=f"Page {self.current_page + 1} / {total_pages}")
        self.prev_button.config(state="normal" if self.current_page > 0 else "disabled")
        self.next_button.config(state="normal" if self.current_page < total_pages - 1 else "disabled")
        self.back_button.config(state="normal" if self.history else "disabled")

    # Рисува home dashboard върху текущия екран.
    def _render_home_dashboard(self) -> None:
        if not self.cards_frame.winfo_manager():
            self.cards_frame.pack(fill="both", expand=True)
        self.cards_frame.lift()
        self._finish_render_home_dashboard()

    # Показва dashboard direct в интерфейса.
    def _show_dashboard_direct(self, reset_history: bool = True) -> None:
        # Force path for every Dashboard button: it always redraws the real dashboard renderer.
        if reset_history:
            self.history.clear()
        self.current_menu = "main"
        self.current_page = 0
        menu = MENU_TREE["main"]
        self.menu_path.config(text=self._build_path())
        self.card_title.config(text=menu["title"])
        self.card_subtitle.config(text=menu["subtitle"])
        self.subtitle_label.config(text=menu["subtitle"])
        self.header_dashboard_button.config(state="normal")
        self.dashboard_button.config(state="disabled")
        self._refresh_sidebar_navigation()
        self._refresh_overview_cards()
        self._toggle_dashboard_chrome(True)
        self._toggle_language_status_panel(False)
        self._finish_render_home_dashboard()
        self.status_var.set("Dashboard е зареден.")
        self.root.after(120, self._ensure_dashboard_visible)

    # Помощна функция за finish render home dashboard.
    def _finish_render_home_dashboard(self) -> None:
        if self.dashboard_is_rendering:
            return
        self.dashboard_is_rendering = True
        self.dashboard_render_job = None
        try:
            self._stop_dashboard_info_scroll()
            self.dashboard_live_widgets = {}
            self.dashboard_host_frame = None
            for widget in self.cards_frame.winfo_children():
                if widget is self.language_status_panel:
                    widget.grid_forget()
                    continue
                widget.destroy()
            for index in range(12):
                self.cards_frame.rowconfigure(index, weight=0, minsize=0)
                self.cards_frame.columnconfigure(index, weight=0, minsize=0)
            self.cards_frame.columnconfigure(0, weight=1)
            self.cards_frame.rowconfigure(0, weight=1)
            dashboard_host = tk.Frame(self.cards_frame, bg=APP_BG, bd=0)
            dashboard_host.pack(fill="both", expand=True)
            self.dashboard_host_frame = dashboard_host

            try:
                self._render_main_dashboard_v2(dashboard_host)
            except Exception as exc:
                self._render_home_dashboard_fallback(exc, traceback.format_exc())
            self.cards_frame.lift()
            self.root.update_idletasks()
        finally:
            self.dashboard_is_rendering = False

    # Помощна функция за ensure dashboard visible.
    def _ensure_dashboard_visible(self) -> None:
        if self.current_menu != "main" or not self.cards_frame.winfo_exists():
            return
        if not self.cards_frame.winfo_manager():
            self.cards_frame.pack(fill="both", expand=True)
        self.root.update_idletasks()
        if self._dashboard_has_visible_content():
            return
        self._finish_render_home_dashboard()
        self.root.after(120, self._verify_dashboard_visible)

    # Помощна функция за verify dashboard visible.
    def _verify_dashboard_visible(self) -> None:
        if self.current_menu != "main" or not self.cards_frame.winfo_exists():
            return
        self.root.update_idletasks()
        if self._dashboard_has_visible_content():
            return
        self._show_dashboard_debug_message(
            "Dashboard не се вижда след принудително прерисуване.",
            self._dashboard_debug_report("no visible dashboard widgets"),
        )

    # Помощна функция за dashboard has visible content.
    def _dashboard_has_visible_content(self) -> bool:
        host = self.dashboard_host_frame
        if isinstance(host, tk.Frame) and host.winfo_exists() and host.winfo_ismapped():
            return bool(host.winfo_children()) and any(
                child.winfo_exists() and child.winfo_ismapped() and child.winfo_width() > 1 and child.winfo_height() > 1
                for child in host.winfo_children()
            )
        return any(
            child is not self.language_status_panel
            and child.winfo_exists()
            and child.winfo_ismapped()
            and child.winfo_width() > 1
            and child.winfo_height() > 1
            for child in self.cards_frame.winfo_children()
        )

    # Помощна функция за dashboard debug report.
    def _dashboard_debug_report(self, reason: str, traceback_text: str = "") -> str:
        try:
            child_lines = []
            for index, child in enumerate(self.cards_frame.winfo_children(), start=1):
                child_lines.append(
                    f"{index}. {child.winfo_class()} mapped={child.winfo_ismapped()} "
                    f"manager={child.winfo_manager()} size={child.winfo_width()}x{child.winfo_height()}"
                )
            children_text = "\n".join(child_lines) if child_lines else "няма children в cards_frame"
            return (
                f"Причина: {reason}\n"
                f"current_menu={self.current_menu}\n"
                f"history={self.history}\n"
                f"cards_frame manager={self.cards_frame.winfo_manager()} "
                f"mapped={self.cards_frame.winfo_ismapped()} "
                f"size={self.cards_frame.winfo_width()}x{self.cards_frame.winfo_height()}\n"
                f"root size={self.root.winfo_width()}x{self.root.winfo_height()}\n\n"
                f"cards_frame children:\n{children_text}\n\n"
                f"{traceback_text}"
            )
        except Exception as report_exc:
            return f"Неуспешно събиране на dashboard диагностика: {report_exc}"

    # Показва dashboard debug message в интерфейса.
    def _show_dashboard_debug_message(self, title: str, details: str) -> None:
        self.status_var.set(title)
        try:
            messagebox.showerror("Dashboard диагностика", f"{title}\n\n{details}", parent=self.root)
        except tk.TclError:
            pass

    # Рисува home dashboard fallback върху текущия екран.
    def _render_home_dashboard_fallback(self, exc: Exception, traceback_text: str = "") -> None:
        self._show_dashboard_debug_message(
            "Dashboard не може да се зареди.",
            self._dashboard_debug_report(str(exc), traceback_text),
        )
        parent = self.dashboard_host_frame if isinstance(self.dashboard_host_frame, tk.Frame) and self.dashboard_host_frame.winfo_exists() else self.cards_frame
        fallback = tk.Frame(
            parent,
            bg=APP_PANEL,
            bd=0,
            highlightthickness=1,
            highlightbackground=APP_DANGER,
        )
        fallback.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        tk.Label(
            fallback,
            text="Dashboard не може да се зареди",
            font=self._font(16, "bold", "Segoe UI Semibold"),
            fg="#ffb0b0",
            bg=APP_PANEL,
        ).pack(anchor="w", padx=18, pady=(18, 6))
        tk.Label(
            fallback,
            text=str(exc),
            font=self._font(10),
            fg=APP_TEXT_SOFT,
            bg=APP_PANEL,
            justify="left",
            wraplength=max(420, self.right_subtitle_wrap),
        ).pack(anchor="w", fill="x", padx=18, pady=(0, 12))
        tk.Button(
            fallback,
            text="Опитай отново",
            command=self.go_home,
            font=self._font(10, "bold", "Segoe UI Semibold"),
            bg=APP_ACCENT_SOFT,
            fg="#f2fff8",
            activebackground="#27a67a",
            activeforeground="#ffffff",
            bd=0,
            padx=16,
            pady=10,
            cursor="hand2",
        ).pack(anchor="w", padx=18, pady=(0, 18))
        self.page_label.config(text="Dashboard error")
        self.prev_button.config(state="disabled")
        self.next_button.config(state="disabled")
        self.back_button.config(state="disabled")

    # Рисува activation menu върху текущия екран.
    def _render_activation_menu(self) -> None:
        self.cards_frame.columnconfigure(0, weight=1)
        self.cards_frame.columnconfigure(1, weight=1)
        self.cards_frame.columnconfigure(2, weight=1)
        self.cards_frame.rowconfigure(0, weight=0)
        self.cards_frame.rowconfigure(1, weight=1)

        header = tk.Frame(
            self.cards_frame,
            bg="#0b211d",
            bd=0,
            highlightthickness=1,
            highlightbackground=APP_BORDER_STRONG,
        )
        header.grid(row=0, column=0, columnspan=3, sticky="ew", padx=8, pady=(0, 12))

        icon_wrap = tk.Frame(header, bg="#102f29", width=self._scale_px(72), height=self._scale_px(72), highlightthickness=1, highlightbackground=APP_ACCENT)
        icon_wrap.pack(side="left", padx=16, pady=14)
        icon_wrap.pack_propagate(False)
        key_icon = self.dashboard_icons.get("key")
        if key_icon is not None:
            tk.Label(icon_wrap, image=key_icon, bg="#102f29").pack(expand=True)
        else:
            tk.Label(icon_wrap, text="KEY", font=self._font(12, "bold", "Segoe UI Semibold"), fg=APP_ACCENT, bg="#102f29").pack(expand=True)

        text_area = tk.Frame(header, bg="#0b211d")
        text_area.pack(side="left", fill="both", expand=True, pady=14)
        tk.Label(
            text_area,
            text="Activation Control Center",
            font=self._font(18, "bold", "Segoe UI Semibold"),
            fg=APP_TEXT,
            bg="#0b211d",
        ).pack(anchor="w")
        tk.Label(
            text_area,
            text="Избери Windows или Office workflow. Ключовете и действията са отделени в чисти модули, за да няма разсейващи панели.",
            font=self._font(10),
            fg=APP_TEXT_SOFT,
            bg="#0b211d",
            wraplength=max(520, self.right_subtitle_wrap),
            justify="left",
        ).pack(anchor="w", pady=(4, 10))

        chips = tk.Frame(text_area, bg="#0b211d")
        chips.pack(anchor="w")
        for label, color in (("Windows 10", APP_ACCENT_BLUE), ("Windows 11", APP_ACCENT), ("Office", APP_WARNING)):
            chip = tk.Label(
                chips,
                text=label,
                font=self._font(8, "bold", "Segoe UI Semibold"),
                fg="#f7fffb",
                bg="#122f2a",
                padx=12,
                pady=5,
                highlightthickness=1,
                highlightbackground=color,
            )
            chip.pack(side="left", padx=(0, 8))

        items = MENU_TREE[self.current_menu]["items"]
        for column, item in enumerate(items):
            card = self._build_card(self.cards_frame, item)
            card.grid(row=1, column=column, sticky="nsew", padx=8, pady=8)

        self.page_label.config(text="Activation Center")
        self.prev_button.config(state="disabled")
        self.next_button.config(state="disabled")
        self.back_button.config(state="normal" if self.history else "disabled")

    # Помощна функция за discover standalone installers.
    def _discover_standalone_installers(self, known_relative_files: set[str]) -> list[dict[str, str]]:
        # Намира локални installer файлове в Installers, които още не са описани ръчно.
        installers_root = self.resource_status.installers_root
        if not installers_root.exists():
            return []

        office_folders = {
            installer.folder.casefold()
            for installer in OFFICE_OFFLINE_INSTALLERS.values()
        }
        discovered: list[dict[str, str]] = []
        runnable_extensions = {".exe", ".msi", ".bat", ".cmd"}
        seen_paths: set[str] = set()

        for child in installers_root.iterdir():
            child_name = child.name.casefold()
            if child_name in office_folders:
                continue
            candidates: list[Path] = []
            if child.is_file() and child.suffix.lower() in runnable_extensions:
                candidates.append(child)
            elif child.is_dir():
                candidates.extend(path for path in child.rglob("*") if path.is_file() and path.suffix.lower() in runnable_extensions)

            for path in candidates:
                relative_path = path.relative_to(installers_root).as_posix()
                if relative_path in known_relative_files or relative_path in seen_paths:
                    continue
                seen_paths.add(relative_path)
                discovered.append(
                    {
                        "id": f"standalone_{relative_path.replace('/', '_').replace(' ', '_').replace('.', '_')}",
                        "label": path.stem,
                        "category": "Локални инструменти",
                        "description": f"Стартира локалния файл: {relative_path}",
                        "type": "standalone_local",
                        "local_path": str(path),
                    }
                )
        return discovered

    # Помощна функция за health item value.
    def _health_item_value(self, label: str) -> tuple[str, bool]:
        # Намира последната известна стойност за конкретен health ред.
        for item in self.latest_health_items:
            if item.label == label:
                return item.value, item.ok
        return "Няма данни", False

    # Помощна функция за dashboard system rows.
    def _dashboard_system_rows(self) -> list[tuple[str, str]]:
        # Подрежда всички налични системни данни за плаващата карта в началото.
        source_items = self.latest_health_items
        if not source_items:
            return [
                ("Общо състояние", "Зареждане..."),
                ("Компютър", os.environ.get("COMPUTERNAME", "Няма данни")),
                ("Потребител", os.environ.get("USERNAME", "Няма данни")),
                ("Операционна система", f"{platform.system()} {platform.release()}".strip() or "Няма данни"),
                ("IP адрес", "Зареждане..."),
                ("Време на работа", "Зареждане..."),
                ("Процесор", platform.processor() or "Зареждане..."),
                ("Натоварване на процесора", "Зареждане..."),
                ("Температура на процесора", "Зареждане..."),
                ("Напрежение на процесора", "Зареждане..."),
                ("RAM използване", "Зареждане..."),
                ("RAM тип и скорост", "Зареждане..."),
                ("Графична карта", "Зареждане..."),
                ("Дънна платка", "Зареждане..."),
                ("BIOS версия", "Зареждане..."),
                ("Secure Boot", "Зареждане..."),
                ("Батерия", "Зареждане..."),
                ("Дискове", "Зареждане..."),
            ]

        health_map = {item.label: item.value for item in source_items}
        computer_user = health_map.get("PC/User:", "Няма данни")
        if " / " in computer_user:
            computer_name, user_name = computer_user.split(" / ", 1)
        else:
            computer_name = os.environ.get("COMPUTERNAME", "Няма данни")
            user_name = os.environ.get("USERNAME", "Няма данни")
        label_map = {
            "OS:": "Операционна система",
            "IP:": "IP адрес",
            "Uptime:": "Време на работа",
            "CPU:": "Процесор",
            "CPU Load:": "Натоварване на процесора",
            "Temperature:": "Температура на процесора",
            "CPU Voltage:": "Напрежение на процесора",
            "GPU:": "Графична карта",
            "RAM:": "RAM използване",
            "RAM Type:": "RAM тип и скорост",
            "Motherboard:": "Дънна платка",
            "BIOS:": "BIOS версия",
            "Secure Boot:": "Secure Boot",
            "Battery:": "Батерия",
        }
        preferred_order = [
            "OS:",
            "PC/User:",
            "IP:",
            "Uptime:",
            "CPU:",
            "CPU Load:",
            "Temperature:",
            "CPU Voltage:",
            "GPU:",
            "RAM:",
            "RAM Type:",
            "Motherboard:",
            "BIOS:",
            "Secure Boot:",
            "Battery:",
        ]

        rows: list[tuple[str, str]] = []
        system_status = "OK" if all(item.ok for item in source_items) and source_items else "Има внимание"
        rows.append(("Общо състояние", system_status))
        rows.append(("Компютър", computer_name))
        rows.append(("Потребител", user_name))

        for key in preferred_order:
            if key == "PC/User:":
                continue
            value = health_map.get(key, "").strip() or "Няма данни"
            rows.append((label_map.get(key, key.rstrip(":")), value))

        disk_found = False
        for item in source_items:
            if item.label.startswith("Disk "):
                disk_found = True
                disk_label = item.label.replace("Disk ", "Диск ").rstrip(":")
                rows.append((disk_label, item.value))
        if not disk_found:
            rows.append(("Дискове", "Няма данни"))

        return rows

    # Помощна функция за dashboard component rows.
    def _dashboard_component_rows(self) -> list[tuple[str, str, bool]]:
        if self.component_status_cache:
            return list(self.component_status_cache)
        # Събира десния списък със статуси на компоненти.
        windows_value, windows_ok = self._component_windows_activation_status()
        office_value, office_ok = self._component_office_activation_status()
        net_value, net_ok = self._component_dotnet_status()
        directx_value, directx_ok = self._component_directx_status()
        vc_value, vc_ok = self._component_visual_cpp_status()
        defender_value, defender_ok = self._component_defender_status()
        firewall_value, firewall_ok = self._component_firewall_status()
        bitlocker_value, bitlocker_ok = self._component_bitlocker_status()
        return [
            ("Windows статус", windows_value, windows_ok),
            ("Office статус", office_value, office_ok),
            (".NET Framework", net_value, net_ok),
            ("DirectX", directx_value, directx_ok),
            ("Visual C++ Redistributable", vc_value, vc_ok),
            ("Windows Defender", defender_value, defender_ok),
            ("Firewall", firewall_value, firewall_ok),
            ("BitLocker", bitlocker_value, bitlocker_ok),
        ]

    # Помощна функция за dashboard component rows for dashboard.
    def _dashboard_component_rows_for_dashboard(self) -> list[tuple[str, str, bool]]:
        if self.component_status_cache:
            return list(self.component_status_cache)
        self._refresh_component_status_async()
        return [
            ("Windows статус", "Проверява се...", False),
            ("Office статус", "Проверява се...", False),
            (".NET Framework", "Проверява се...", False),
            ("DirectX", "Проверява се...", False),
            ("Visual C++ Redistributable", "Проверява се...", False),
            ("Windows Defender", "Проверява се...", False),
            ("Firewall", "Проверява се...", False),
            ("BitLocker", "Проверява се...", False),
        ]

    # Помощна функция за refresh component status async.
    def _refresh_component_status_async(self) -> None:
        if self.component_status_refresh_in_progress:
            return
        self.component_status_refresh_in_progress = True

        # Помощна функция за worker.
        def worker() -> None:
            try:
                rows = self._dashboard_component_rows()
            except Exception as exc:
                rows = [("Компонентен статус", f"Грешка: {exc}", False)]
            try:
                self.root.after(0, lambda: self._apply_component_status_rows(rows))
            except RuntimeError:
                self.component_status_refresh_in_progress = False

        threading.Thread(target=worker, daemon=True).start()

    # Помощна функция за apply component status rows.
    def _apply_component_status_rows(self, rows: list[tuple[str, str, bool]]) -> None:
        self.component_status_refresh_in_progress = False
        self.component_status_cache = list(rows)
        if self.current_menu == "main":
            self._render_cards()

    # Помощна функция за component windows activation status.
    def _component_windows_activation_status(self) -> tuple[str, bool]:
        # Проверява реалния статус на активацията на Windows.
        script = (
            "$product = Get-CimInstance SoftwareLicensingProduct | "
            "Where-Object { $_.PartialProductKey -and $_.ApplicationID -eq "
            "'55c92734-d682-4d71-983e-d6ec3f16059f' } | Select-Object -First 1; "
            "if ($null -eq $product) { 'NONE' } else { "
            "\"$($product.LicenseStatus)|$($product.Description)|$($product.PartialProductKey)\" }"
        )
        try:
            result = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
                capture_output=True,
                text=True,
                timeout=12,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception:
            saved_windows_key = any(
                bool(self.secure_store.get(key, "").strip())
                for key in ("windows10_product_key", "windows11_product_key")
            )
            return ("Има записан ключ", False) if saved_windows_key else ("Няма данни", False)
        output = "\n".join(part.strip() for part in (result.stdout, result.stderr) if part and part.strip())
        if "0x80041003" in output or "access denied" in output.lower():
            saved_windows_key = any(
                bool(self.secure_store.get(key, "").strip())
                for key in ("windows10_product_key", "windows11_product_key")
            )
            return ("Провери като админ", False) if not saved_windows_key else ("Има записан ключ", False)
        if not output or output.strip() == "NONE":
            return "Няма данни", False
        first_line = output.splitlines()[0].strip()
        parts = [part.strip() for part in first_line.split("|")]
        try:
            license_status = int(parts[0])
        except Exception:
            license_status = 0
        description = parts[1] if len(parts) > 1 else ""
        partial_key = parts[2] if len(parts) > 2 else ""
        if license_status == 1:
            tail = f" ({partial_key})" if partial_key else ""
            return f"Активиран{tail}", True
        if "notification" in description.lower():
            return "Иска активация", False
        if partial_key:
            return f"Има ключ ({partial_key})", False
        saved_windows_key = any(
            bool(self.secure_store.get(key, "").strip())
            for key in ("windows10_product_key", "windows11_product_key")
        )
        return ("Има записан ключ", False) if saved_windows_key else ("Неактивиран", False)

    # Помощна функция за component office activation status.
    def _component_office_activation_status(self) -> tuple[str, bool]:
        # Проверява реалния статус на Office през OSPP.VBS.
        ospp_vbs = find_ospp_vbs()
        if not ospp_vbs:
            saved_office_key = any(bool(self.secure_store.get(f"{key}_product_key", "").strip()) for key in OFFICE_ACTION_IDS)
            return ("Има записан Office ключ", False) if saved_office_key else ("Office не е открит", False)
        try:
            result = subprocess.run(
                ["cscript", "//nologo", str(ospp_vbs), "/dstatus"],
                capture_output=True,
                text=True,
                timeout=20,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception:
            saved_office_key = any(bool(self.secure_store.get(f"{key}_product_key", "").strip()) for key in OFFICE_ACTION_IDS)
            return ("Има записан Office ключ", False) if saved_office_key else ("Няма данни", False)
        output = "\n".join(part.strip() for part in (result.stdout, result.stderr) if part and part.strip())
        if not output:
            return "Няма данни", False
        upper_output = output.upper()
        if "---LICENSED---" in upper_output or "LICENSE STATUS:  ---LICENSED---" in upper_output:
            match = re.search(r"Last 5 characters of installed product key:\s*([A-Z0-9]{5})", output, re.IGNORECASE)
            tail = f" ({match.group(1)})" if match else ""
            return f"Активиран{tail}", True
        if "0X1A8" in upper_output or "ACCESS DENIED" in upper_output:
            saved_office_key = any(bool(self.secure_store.get(f"{key}_product_key", "").strip()) for key in OFFICE_ACTION_IDS)
            return ("Провери като админ", False) if not saved_office_key else ("Има записан Office ключ", False)
        if "LICENSE STATUS" in upper_output:
            return "Има Office, но не е активиран", False
        saved_office_key = any(bool(self.secure_store.get(f"{key}_product_key", "").strip()) for key in OFFICE_ACTION_IDS)
        return ("Има записан Office ключ", False) if saved_office_key else ("Няма данни", False)

    # Помощна функция за component service running.
    def _component_service_running(self, service_name: str) -> bool | None:
        # Проверява дали дадена Windows услуга е стартирана.
        try:
            result = subprocess.run(
                ["sc", "query", service_name],
                capture_output=True,
                text=True,
                timeout=8,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception:
            return None
        output = f"{result.stdout}\n{result.stderr}".upper()
        if "STATE" not in output:
            return None
        if "RUNNING" in output:
            return True
        if "STOPPED" in output:
            return False
        return None

    # Помощна функция за component dotnet status.
    def _component_dotnet_status(self) -> tuple[str, bool]:
        # Чете .NET 4.x версията от registry.
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full")
            release = int(winreg.QueryValueEx(key, "Release")[0])
        except Exception:
            return "Няма данни", False
        if release >= 533320:
            return ".NET 4.8.1", True
        if release >= 528040:
            return ".NET 4.8", True
        if release >= 461808:
            return ".NET 4.7.2", True
        return f"Release {release}", True

    # Помощна функция за component directx status.
    def _component_directx_status(self) -> tuple[str, bool]:
        # Проверява наличието на DirectX 12 runtime по системния DLL.
        system_root = Path(os.environ.get("WINDIR", r"C:\Windows"))
        d3d12_path = system_root / "System32" / "d3d12.dll"
        if d3d12_path.exists():
            return "DirectX 12 наличен", True
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\DirectX")
            version = str(winreg.QueryValueEx(key, "Version")[0]).strip()
            if version:
                return version, True
        except Exception:
            pass
        return "Няма данни", False

    # Помощна функция за component visual cpp status.
    def _component_visual_cpp_status(self) -> tuple[str, bool]:
        # Търси инсталирани Visual C++ пакети в uninstall списъка.
        found: list[str] = []
        uninstall_paths = (
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        )
        for uninstall_path in uninstall_paths:
            try:
                root_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, uninstall_path)
            except OSError:
                continue
            index = 0
            while True:
                try:
                    subkey_name = winreg.EnumKey(root_key, index)
                    index += 1
                except OSError:
                    break
                try:
                    subkey = winreg.OpenKey(root_key, subkey_name)
                    display_name = str(winreg.QueryValueEx(subkey, "DisplayName")[0]).strip()
                except OSError:
                    continue
                if "Visual C++" in display_name:
                    found.append(display_name)
        if not found:
            return "Липсват", False
        return f"Налични ({len(found)})", True

    # Помощна функция за component defender status.
    def _component_defender_status(self) -> tuple[str, bool]:
        # Проверява дали услугата на Windows Defender работи.
        running = self._component_service_running("WinDefend")
        if running is None:
            return "Няма данни", False
        return ("Активен", True) if running else ("Спрян", False)

    # Помощна функция за component firewall status.
    def _component_firewall_status(self) -> tuple[str, bool]:
        # Проверява дали услугата на Windows Firewall работи.
        running = self._component_service_running("MpsSvc")
        if running is None:
            return "Няма данни", False
        return ("Активна", True) if running else ("Изключена", False)

    # Помощна функция за component bitlocker status.
    def _component_bitlocker_status(self) -> tuple[str, bool]:
        # Проверява BitLocker чрез manage-bde и пада към по-лек service fallback.
        try:
            result = subprocess.run(
                ["manage-bde", "-status", "C:"],
                capture_output=True,
                text=True,
                timeout=12,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            output = f"{result.stdout}\n{result.stderr}".lower()
            if "protection on" in output:
                return "Активен", True
            if "protection off" in output or "fully decrypted" in output:
                return "Изключен", False
        except Exception:
            pass
        running = self._component_service_running("BDESVC")
        if running is None:
            return "Няма данни", False
        return ("Готов", True) if running else ("Изключен", False)

    # Помощна функция за dashboard metric cards.
    def _dashboard_metric_cards(self) -> list[dict[str, object]]:
        # Горните бързи метрики по снимката.
        temp_value, temp_ok = self._health_item_value("Temperature:")
        voltage_value, voltage_ok = self._health_item_value("CPU Voltage:")
        ram_value, ram_ok = self._health_item_value("RAM:")
        ram_type_value, _ = self._health_item_value("RAM Type:")
        disk_value, disk_ok = self._health_item_value("Disk C:")
        ram_parts = [part.strip() for part in ram_type_value.split("/") if part.strip()]
        ram_type_short = ram_parts[1] if len(ram_parts) > 1 else (ram_parts[0] if ram_parts else "Unknown")
        ram_speed_short = ram_parts[2] if len(ram_parts) > 2 else "Speed N/A"
        return [
            {"key": "cpu_temp", "icon": "cpu", "title": "CPU температура", "value": temp_value, "status": "Нормална" if temp_ok else "Провери", "ok": temp_ok},
            {"key": "cpu_voltage", "icon": "bolt", "title": "CPU напрежение", "value": voltage_value, "status": "Стабилно" if voltage_ok else "Няма данни", "ok": voltage_ok},
            {"key": "ram_usage", "icon": "ram", "title": "RAM използване", "value": ram_value, "status": f"{ram_type_short} • {ram_speed_short}", "ok": ram_ok},
            {"key": "disk_c", "icon": "disk", "title": "Системен диск (C:)", "value": disk_value, "status": "Добро" if disk_ok else "Запълнен", "ok": disk_ok},
        ]

    # Помощна функция за dashboard metric percent.
    def _dashboard_metric_percent(self, value: str, fallback_ok: bool) -> int:
        # Изчислява процента за живите dashboard карти.
        percent_match = re.search(r"(\d{1,3})\s*%", value)
        if percent_match:
            return max(0, min(100, int(percent_match.group(1))))
        ratio_match = re.search(r"([0-9]+(?:[.,][0-9]+)?)\s*(?:GB|V|В°C)?.*?/\s*([0-9]+(?:[.,][0-9]+)?)", value)
        if ratio_match:
            try:
                used = float(ratio_match.group(1).replace(",", "."))
                total = float(ratio_match.group(2).replace(",", "."))
                if total > 0:
                    return max(0, min(100, int((used / total) * 100)))
            except ValueError:
                pass
        return 72 if fallback_ok else 26

    # Помощна функция за dashboard quick actions.
    def _dashboard_quick_actions(self) -> list[dict[str, str]]:
        # Бързи действия в долния десен панел.
        return [
            {"icon": "actions_small", "label": "Почистване\nна системата", "action_id": "reset_onedrive_2"},
            {"icon": "refresh_small", "label": "Рестартиране\nна услуги", "action_id": "reset_onedrive_1"},
            {"icon": "monitor_small", "label": "Проверка\nна здравето", "action_id": "driver_pc_report"},
            {"icon": "globe_small", "label": "Езици\nКлавиатури", "menu": "language"},
        ]

    # Рисува main dashboard върху текущия екран.
    def _render_main_dashboard(self) -> None:
        # Специален dashboard renderer за Начало по новия дизайн.
        self.cards_frame.columnconfigure(0, weight=1)
        self.cards_frame.rowconfigure(0, weight=1)

        outer = tk.Frame(
            self.cards_frame,
            bg=APP_BG,
            bd=0,
        )
        outer.grid(row=0, column=0, sticky="nsew")

        status_ok = all(item.ok for item in self.latest_health_items) if self.latest_health_items else False
        banner = tk.Frame(outer, bg=APP_PANEL_SOFT, bd=0, highlightthickness=1, highlightbackground=APP_BORDER_STRONG)
        banner.pack(fill="x", pady=(0, 10))
        banner_icon = self.dashboard_icons.get("shield")
        if banner_icon is not None:
            tk.Label(banner, image=banner_icon, bg=APP_PANEL_SOFT).pack(side="left", padx=(18, 14), pady=16)
        else:
            tk.Label(banner, text="✓", font=self._font(22, "bold", "Segoe UI Semibold"), fg=APP_ACCENT, bg=APP_PANEL_SOFT).pack(side="left", padx=(18, 14), pady=16)
        banner_text = tk.Frame(banner, bg=APP_PANEL_SOFT)
        banner_text.pack(side="left", fill="x", expand=True, pady=14)
        tk.Label(
            banner_text,
            text="Системата е защитена и работи оптимално!" if status_ok else "Има компоненти, които искат внимание",
            font=self._font(16, "bold", "Segoe UI Semibold"),
            fg=APP_ACCENT if status_ok else "#ff8b8b",
            bg=APP_PANEL_SOFT,
        ).pack(anchor="w")
        tk.Label(
            banner_text,
            text="Всички критични компоненти са в добро състояние." if status_ok else "Провери червените статуси и препоръчаните действия вдясно.",
            font=self._font(10),
            fg=APP_TEXT_SOFT,
            bg=APP_PANEL_SOFT,
        ).pack(anchor="w", pady=(4, 0))
        banner_right = tk.Frame(banner, bg=APP_PANEL_SOFT)
        banner_right.pack(side="right", padx=18, pady=12)
        update_text = self.update_message_var.get() if hasattr(self, "update_message_var") else "Няма данни за update"
        tk.Label(banner_right, text="Статус на ъпдейти", font=self._font(9), fg=APP_TEXT_MUTED, bg=APP_PANEL_SOFT).pack(anchor="w")
        tk.Label(banner_right, text=update_text, font=self._font(11, "bold", "Segoe UI Semibold"), fg=APP_TEXT, bg=APP_PANEL_SOFT, wraplength=360, justify="left").pack(anchor="w", pady=(2, 8))
        tk.Button(
            banner_right,
            text="Провери отново",
            command=self._check_updates_async,
            font=self._font(9, "bold", "Segoe UI Semibold"),
            bg=APP_ACCENT_SOFT,
            fg="#f2fff8",
            activebackground="#27a67a",
            activeforeground="#ffffff",
            bd=0,
            padx=18,
            pady=8,
            cursor="hand2",
        ).pack(anchor="e")

        metrics = tk.Frame(outer, bg=APP_BG)
        metrics.pack(fill="x", pady=(0, 12))
        for index in range(5):
            metrics.columnconfigure(index, weight=1, uniform="metrics")
        for idx, spec in enumerate(self._dashboard_metric_cards()):
            card = tk.Frame(metrics, bg=APP_PANEL, bd=0, highlightthickness=1, highlightbackground=APP_BORDER)
            card.grid(row=0, column=idx, sticky="nsew", padx=(0 if idx == 0 else 8, 0))
            icon_image = self.dashboard_icons.get(str(spec["icon"]))
            if icon_image is not None:
                tk.Label(card, image=icon_image, bg=APP_PANEL).pack(anchor="w", padx=16, pady=(12, 0))
            else:
                tk.Label(card, text=str(spec["icon"]), font=self._font(20, "bold", "Segoe UI Symbol"), fg=APP_ACCENT, bg=APP_PANEL).pack(anchor="w", padx=16, pady=(12, 0))
            tk.Label(card, text=str(spec["title"]), font=self._font(10), fg=APP_TEXT_SOFT, bg=APP_PANEL).pack(anchor="w", padx=16, pady=(6, 0))
            tk.Label(card, text=str(spec["value"]), font=self._font(18, "bold", "Segoe UI Semibold"), fg=APP_TEXT if spec["ok"] else "#ff8b8b", bg=APP_PANEL).pack(anchor="w", padx=16, pady=(6, 0))
            tk.Label(card, text=str(spec["status"]), font=self._font(9), fg=APP_ACCENT if spec["ok"] else "#ff8b8b", bg=APP_PANEL).pack(anchor="w", padx=16, pady=(6, 12))

        alert_card = tk.Frame(metrics, bg="#2a1719", bd=0, highlightthickness=1, highlightbackground="#5a2a30")
        alert_card.grid(row=0, column=4, sticky="nsew", padx=(8, 0))
        problem_count = sum(1 for item in self.latest_health_items if not item.ok)
        warning_icon = self.dashboard_icons.get("warning")
        if warning_icon is not None:
            tk.Label(alert_card, image=warning_icon, bg="#2a1719").pack(anchor="w", padx=16, pady=(12, 0))
        else:
            tk.Label(alert_card, text="!", font=self._font(20, "bold", "Segoe UI Semibold"), fg="#ff6b6b", bg="#2a1719").pack(anchor="w", padx=16, pady=(12, 0))
        tk.Label(alert_card, text="Проверка на сигурност", font=self._font(10), fg="#f3b1b1", bg="#2a1719").pack(anchor="w", padx=16, pady=(6, 0))
        tk.Label(alert_card, text="Внимание" if problem_count else "Няма проблем", font=self._font(17, "bold", "Segoe UI Semibold"), fg="#ff6b6b" if problem_count else APP_ACCENT, bg="#2a1719").pack(anchor="w", padx=16, pady=(6, 0))
        tk.Label(alert_card, text=f"{problem_count} проблем(а) открити" if problem_count else "Системата изглежда стабилна", font=self._font(9), fg="#ffd1d1" if problem_count else APP_TEXT_SOFT, bg="#2a1719").pack(anchor="w", padx=16, pady=(6, 12))

        lower = tk.Frame(outer, bg=APP_BG)
        lower.pack(fill="both", expand=True)
        lower.columnconfigure(0, weight=10, uniform="lower")
        lower.columnconfigure(1, weight=14, uniform="lower")
        lower.columnconfigure(2, weight=12, uniform="lower")

        info_panel = tk.Frame(lower, bg=APP_PANEL, bd=0, highlightthickness=1, highlightbackground=APP_BORDER)
        info_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        info_header = tk.Frame(info_panel, bg=APP_PANEL)
        info_header.pack(fill="x", padx=16, pady=(14, 10))
        if self.dashboard_icons.get("monitor_small") is not None:
            tk.Label(info_header, image=self.dashboard_icons["monitor_small"], bg=APP_PANEL).pack(side="left", padx=(0, 8))
        tk.Label(info_header, text="Информация за системата", font=self._font(14, "bold", "Segoe UI Semibold"), fg=APP_TEXT, bg=APP_PANEL).pack(side="left")
        for label, value in self._dashboard_system_rows():
            row = tk.Frame(info_panel, bg=APP_PANEL_ALT)
            row.pack(fill="x", padx=14, pady=4)
            tk.Label(row, text=f"{label}:", font=self._font(9), fg=APP_TEXT_MUTED, bg=APP_PANEL_ALT, width=18, anchor="w").pack(side="left", padx=10, pady=8)
            tk.Label(row, text=value, font=self._font(9), fg=APP_TEXT, bg=APP_PANEL_ALT, justify="left", anchor="w", wraplength=280).pack(side="left", fill="x", expand=True, padx=(0, 10), pady=8)
        tk.Button(
            info_panel_body,
            text="Подробен системен отчет",
            command=lambda: self._handle_driver_backup_action("driver_pc_report"),
            font=self._font(9, "bold", "Segoe UI Semibold"),
            bg=APP_PANEL_ALT,
            fg=APP_TEXT,
            activebackground=APP_BORDER_STRONG,
            activeforeground="#ffffff",
            bd=0,
            padx=16,
            pady=10,
            cursor="hand2",
        ).pack(fill="x", padx=14, pady=(14, 14))

        installer_panel = tk.Frame(lower, bg=APP_PANEL, bd=0, highlightthickness=1, highlightbackground=APP_BORDER)
        installer_panel.grid(row=0, column=1, sticky="nsew", padx=8)
        header_row = tk.Frame(installer_panel, bg=APP_PANEL)
        header_row.pack(fill="x", padx=16, pady=(14, 10))
        if self.dashboard_icons.get("robot_small") is not None:
            tk.Label(header_row, image=self.dashboard_icons["robot_small"], bg=APP_PANEL).pack(side="left", padx=(0, 8))
        tk.Label(header_row, text="Автоматичен инсталатор", font=self._font(14, "bold", "Segoe UI Semibold"), fg=APP_TEXT, bg=APP_PANEL).pack(side="left")
        tk.Button(
            header_row,
            text="Управление на задачи",
            command=lambda: self.render_menu("auto_installer"),
            font=self._font(8, "bold", "Segoe UI Semibold"),
            bg=APP_ACCENT_SOFT,
            fg="#f2fff8",
            activebackground="#27a67a",
            activeforeground="#ffffff",
            bd=0,
            padx=14,
            pady=7,
            cursor="hand2",
        ).pack(side="right")
        tk.Label(installer_panel, text="Изберете какво да инсталирате", font=self._font(9), fg=APP_TEXT_SOFT, bg=APP_PANEL).pack(anchor="w", padx=16)
        list_holder = tk.Frame(installer_panel, bg=APP_PANEL_ALT)
        list_holder.pack(fill="both", expand=True, padx=14, pady=10)
        preview_tasks = self._auto_install_tasks()
        self._ensure_auto_install_vars(preview_tasks)
        dashboard_count_var = tk.StringVar(value="Избрани задачи: 0")

        # Обновява dashboard install count след промяна в състоянието.
        def update_dashboard_install_count(*_: object) -> None:
            selected_count = sum(
                1
                for task in preview_tasks
                if self.auto_install_vars.get(task["id"]) and self.auto_install_vars[task["id"]].get()
            )
            dashboard_count_var.set(f"Избрани задачи: {selected_count}")

        # Задава dashboard install selection според избраното действие.
        def set_dashboard_install_selection(value: bool) -> None:
            for task in preview_tasks:
                task_id = task["id"]
                installed_now, _detail = self._dashboard_task_install_state(task)
                if installed_now:
                    self.auto_install_vars[task_id].set(False)
                    continue
                self.auto_install_vars[task_id].set(value)
            update_dashboard_install_count()
        category = ""
        selected_count = 0
        for index, task in enumerate(preview_tasks, start=1):
            if task["category"] != category:
                category = task["category"]
                tk.Label(list_holder, text=category, font=self._font(10, "bold", "Segoe UI Semibold"), fg=APP_ACCENT, bg=APP_PANEL_ALT).pack(anchor="w", padx=12, pady=(10, 4))
            installed_now, installed_text = self._dashboard_task_install_state(task)
            row = tk.Frame(list_holder, bg=APP_PANEL_ALT)
            row.pack(fill="x", padx=12, pady=2)
            tk.Label(row, text=f"□ {index}. {task['label']}", font=self._font(9), fg=APP_TEXT, bg=APP_PANEL_ALT, anchor="w").pack(side="left", fill="x", expand=True)
            tk.Label(row, text="Наличен" if installed_now else "Липсва", font=self._font(8), fg=APP_ACCENT if installed_now else APP_WARNING, bg=APP_PANEL_ALT).pack(side="right")
        footer = tk.Frame(installer_panel, bg=APP_PANEL)
        footer.pack(fill="x", padx=14, pady=(0, 14))
        tk.Label(footer, text=f"Избрани задачи: {selected_count}", font=self._font(9), fg=APP_TEXT_SOFT, bg=APP_PANEL).pack(side="left")
        tk.Button(
            footer,
            text="Избери всичко",
            command=lambda: self.render_menu("auto_installer"),
            font=self._font(8, "bold", "Segoe UI Semibold"),
            bg=APP_PANEL_ALT,
            fg=APP_TEXT,
            activebackground=APP_BORDER_STRONG,
            activeforeground="#ffffff",
            bd=0,
            padx=12,
            pady=6,
            cursor="hand2",
        ).pack(side="right", padx=(8, 0))
        tk.Button(
            installer_panel_body,
            text="Стартирай инсталацията",
            command=lambda: self.render_menu("auto_installer"),
            font=self._font(10, "bold", "Segoe UI Semibold"),
            bg=APP_ACCENT_SOFT,
            fg="#f2fff8",
            activebackground="#27a67a",
            activeforeground="#ffffff",
            bd=0,
            padx=14,
            pady=12,
            cursor="hand2",
        ).pack(fill="x", padx=14, pady=(0, 14))

        right_column = tk.Frame(lower, bg=APP_BG)
        right_column.grid(row=0, column=2, sticky="nsew", padx=(8, 0))
        right_column.rowconfigure(0, weight=3)
        right_column.rowconfigure(1, weight=2)
        component_panel = tk.Frame(right_column, bg=APP_PANEL, bd=0, highlightthickness=1, highlightbackground=APP_BORDER)
        component_panel.grid(row=0, column=0, sticky="nsew", pady=(0, 8))
        component_header = tk.Frame(component_panel, bg=APP_PANEL)
        component_header.pack(fill="x", padx=16, pady=(14, 10))
        if self.dashboard_icons.get("shield_small") is not None:
            tk.Label(component_header, image=self.dashboard_icons["shield_small"], bg=APP_PANEL).pack(side="left", padx=(0, 8))
        tk.Label(component_header, text="Състояние на компонентите", font=self._font(14, "bold", "Segoe UI Semibold"), fg=APP_TEXT, bg=APP_PANEL).pack(side="left")
        for label, value, ok in self._dashboard_component_rows_for_dashboard():
            row = tk.Frame(component_panel, bg=APP_PANEL_ALT)
            row.pack(fill="x", padx=14, pady=4)
            tk.Label(row, text=label, font=self._font(9), fg=APP_TEXT, bg=APP_PANEL_ALT, anchor="w").pack(side="left", padx=10, pady=8)
            tk.Label(row, text=value, font=self._font(9), fg=APP_ACCENT if ok else "#ff6b6b", bg=APP_PANEL_ALT, anchor="e").pack(side="right", padx=10, pady=8)

        quick_panel = tk.Frame(right_column, bg=APP_PANEL, bd=0, highlightthickness=1, highlightbackground=APP_BORDER)
        quick_panel.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        quick_header = tk.Frame(quick_panel, bg=APP_PANEL)
        quick_header.pack(fill="x", padx=16, pady=(14, 10))
        if self.dashboard_icons.get("actions_small") is not None:
            tk.Label(quick_header, image=self.dashboard_icons["actions_small"], bg=APP_PANEL).pack(side="left", padx=(0, 8))
        tk.Label(quick_header, text="Бързи действия", font=self._font(14, "bold", "Segoe UI Semibold"), fg=APP_TEXT, bg=APP_PANEL).pack(side="left")
        actions_frame = tk.Frame(quick_panel, bg=APP_PANEL)
        actions_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        for idx in range(4):
            actions_frame.columnconfigure(idx, weight=1, uniform="quick")
        for idx, action in enumerate(self._dashboard_quick_actions()):
            card = tk.Frame(actions_frame, bg=APP_PANEL_ALT, bd=0, highlightthickness=1, highlightbackground=APP_BORDER)
            card.grid(row=0, column=idx, sticky="nsew", padx=(0 if idx == 0 else 6, 0))
            if "menu" in action:
                command = lambda menu_key=action["menu"]: self.render_menu(menu_key)
            elif action["action_id"] == "open_console":
                command = lambda: messagebox.showinfo("Конзола", "Тази бърза конзола ще бъде вързана в следващата стъпка.", parent=self.root)
            else:
                command = lambda item={"kind": "action", "action_id": action["action_id"], "label": action["label"]}: self._handle_action(item)
            quick_icon = self.dashboard_icons.get(action["icon"])
            text_content = action["label"] if quick_icon is not None else f"{action['icon']}\n\n{action['label']}"
            button = tk.Button(
                card,
                text=text_content,
                command=command,
                font=self._font(9, "bold", "Segoe UI Semibold"),
                bg=APP_PANEL_ALT,
                fg=APP_TEXT,
                activebackground=APP_BORDER_STRONG,
                activeforeground="#ffffff",
                bd=0,
                padx=10,
                pady=14,
                cursor="hand2",
                justify="center",
                wraplength=110,
                image=quick_icon,
                compound="top",
            )
            button.pack(fill="both", expand=True)

        footer = tk.Frame(outer, bg=APP_PANEL, bd=0, highlightthickness=1, highlightbackground=APP_BORDER)
        footer.pack(fill="x", pady=(12, 0))
        disk_c_value, disk_ok = self._health_item_value("Disk C:")
        footer_items = [
            ("Сигурност", "Добра" if status_ok else "Риск", status_ok),
            ("Производителност", "Отлична" if status_ok else "Провери", status_ok),
            ("Дисково състояние", "Добро" if disk_ok else "Внимание", disk_ok),
        ]
        for idx, (title, value, ok) in enumerate(footer_items):
            segment = tk.Frame(footer, bg=APP_PANEL)
            segment.pack(side="left", fill="x", expand=True, padx=14, pady=10)
            tk.Label(segment, text=f"{title}: ", font=self._font(10), fg=APP_TEXT_SOFT, bg=APP_PANEL).pack(side="left")
            tk.Label(segment, text=value, font=self._font(10, "bold", "Segoe UI Semibold"), fg=APP_ACCENT if ok else "#ff8b8b", bg=APP_PANEL).pack(side="left")

        self.page_label.config(text="Начален dashboard")
        self.prev_button.config(state="disabled")
        self.next_button.config(state="disabled")
        self.back_button.config(state="disabled")

    # Рисува main dashboard v2 върху текущия екран.
    def _render_main_dashboard_v2(self, parent: tk.Widget | None = None) -> None:
        # Нов dashboard за Начало, подреден максимално близо до референтната визия.
        dashboard_parent = parent or self.cards_frame

        # Помощна функция за make panel.
        def make_panel(parent: tk.Widget, *, bg: str = APP_PANEL, border: str = APP_BORDER, radius: int = 18) -> tk.Frame:
            return self._build_soft_panel(parent, panel_bg=bg, border=border, radius=radius, base_bg=APP_BG)

        # Помощна функция за make progress.
        def make_progress(parent: tk.Widget, value: int, accent: str = APP_ACCENT) -> tuple[tk.Frame, tk.Frame]:
            track = tk.Frame(parent, bg="#13201c", height=10)
            track.pack(fill="x", pady=(8, 0))
            track.pack_propagate(False)
            fill = tk.Frame(track, bg=accent, width=max(10, int(220 * (value / 100))), height=10)
            fill.place(x=0, y=0)
            return track, fill

        # Помощна функция за icon or text.
        def icon_or_text(parent: tk.Widget, icon_name: str, fallback: str, *, bg: str, fg: str, side: str = "left") -> None:
            icon_image = self.dashboard_icons.get(icon_name)
            if icon_image is not None:
                tk.Label(parent, image=icon_image, bg=bg).pack(side=side, padx=(0, 10))
            else:
                tk.Label(parent, text=fallback, font=self._font(18, "bold", "Segoe UI Symbol"), fg=fg, bg=bg).pack(side=side, padx=(0, 10))

        outer = tk.Frame(dashboard_parent, bg=APP_BG, bd=0)
        outer.pack(fill="both", expand=True)
        self.dashboard_live_widgets = {}

        status_ok = all(item.ok for item in self.latest_health_items) if self.latest_health_items else False
        update_text = self.update_message_var.get() if hasattr(self, "update_message_var") else "Няма данни за ъпдейти"
        problem_count = sum(1 for item in self.latest_health_items if not item.ok)

        banner = make_panel(outer, bg=APP_PANEL_SOFT, border=APP_BORDER_STRONG, radius=22)
        banner.pack(fill="x", pady=(0, 8))
        banner_body = banner.content  # type: ignore[attr-defined]
        banner_strip = tk.Frame(banner_body, bg=APP_ACCENT, height=4)
        banner_strip.pack(fill="x", padx=10, pady=(8, 0))
        banner_inner = tk.Frame(
            banner_body,
            bg=APP_PANEL,
            bd=0,
            highlightthickness=1,
            highlightbackground="#20352d",
        )
        banner_inner.pack(fill="x", padx=10, pady=(6, 8))
        left_banner = tk.Frame(banner_inner, bg=APP_PANEL)
        left_banner.pack(side="left", padx=(12, 10), pady=7)
        icon_or_text(left_banner, "shield", "✓", bg=APP_PANEL, fg=APP_ACCENT)

        mid_banner = tk.Frame(banner_inner, bg=APP_PANEL)
        mid_banner.pack(side="left", fill="x", expand=True, pady=7)
        tk.Label(
            mid_banner,
            text="Системата е защитена и работи оптимално!" if status_ok else "Има компоненти, които искат внимание",
            font=self._font(11, "bold", "Segoe UI Semibold"),
            fg=APP_ACCENT if status_ok else "#ff8b8b",
            bg=APP_PANEL,
        ).pack(anchor="w")
        tk.Label(
            mid_banner,
            text="Всички критични компоненти са в добро състояние." if status_ok else "Провери червените статуси и препоръчаните действия вдясно.",
            font=self._font(7),
            fg=APP_TEXT_SOFT,
            bg=APP_PANEL,
        ).pack(anchor="w", pady=(2, 0))

        right_banner = tk.Frame(banner_inner, bg=APP_PANEL)
        right_banner.pack(side="right", padx=(8, 12), pady=7)
        tk.Label(right_banner, text="Статус на ъпдейти", font=self._font(7), fg=APP_TEXT_MUTED, bg=APP_PANEL).pack(anchor="w")
        tk.Label(
            right_banner,
            text=update_text,
            font=self._font(9, "bold", "Segoe UI Semibold"),
            fg=APP_TEXT,
            bg=APP_PANEL,
            justify="left",
            wraplength=280,
        ).pack(anchor="w", pady=(2, 3))
        tk.Label(
            right_banner,
            text="Приложението е актуално" if "нова версия" not in update_text.lower() else "Има наличен нов пакет",
            font=self._font(6),
            fg=APP_TEXT_SOFT,
            bg=APP_PANEL,
        ).pack(anchor="w")
        tk.Button(
            right_banner,
            text="Провери отново  ↻",
            command=self._check_updates_async,
            font=self._font(8, "bold", "Segoe UI Semibold"),
            bg=APP_ACCENT_SOFT,
            fg="#f2fff8",
            activebackground="#27a67a",
            activeforeground="#ffffff",
            bd=0,
            padx=12,
            pady=5,
            cursor="hand2",
        ).pack(anchor="w", pady=(6, 0))

        metrics = tk.Frame(outer, bg=APP_BG)
        metrics.pack(fill="x", pady=(0, 12))
        for index in range(5):
            metrics.columnconfigure(index, weight=1, uniform="metrics")

        for idx, spec in enumerate(self._dashboard_metric_cards()):
            card = make_panel(metrics, radius=20)
            card.grid(row=0, column=idx, sticky="nsew", padx=(0 if idx == 0 else 8, 0))
            card_body = card.content  # type: ignore[attr-defined]
            accent_color = APP_ACCENT if spec["ok"] else APP_WARNING
            accent_strip = tk.Frame(card_body, bg=accent_color, height=4)
            accent_strip.pack(fill="x", padx=12, pady=(10, 0))
            inner_frame = tk.Frame(
                card_body,
                bg=APP_PANEL,
                bd=0,
                highlightthickness=1,
                highlightbackground="#20352d",
            )
            inner_frame.pack(fill="both", expand=True, padx=12, pady=(8, 10))
            top_row = tk.Frame(inner_frame, bg=APP_PANEL)
            top_row.pack(fill="x", padx=12, pady=(10, 4))
            icon_or_text(top_row, str(spec["icon"]), "•", bg=APP_PANEL, fg=APP_ACCENT)
            tk.Label(top_row, text=str(spec["title"]), font=self._font(7, "bold", "Segoe UI Semibold"), fg=APP_TEXT, bg=APP_PANEL).pack(side="left")
            value_label = tk.Label(
                inner_frame,
                text=str(spec["value"]),
                font=self._font(12, "bold", "Segoe UI Semibold"),
                fg=APP_ACCENT if spec["ok"] else "#ff8b8b",
                bg=APP_PANEL,
            )
            value_label.pack(anchor="w", padx=12, pady=(0, 5))
            bottom_row = tk.Frame(inner_frame, bg=APP_PANEL)
            bottom_row.pack(fill="x", padx=12, pady=(0, 8))
            tk.Label(bottom_row, text="●", font=self._font(10), fg=APP_ACCENT if spec["ok"] else "#ff8b8b", bg=APP_PANEL).pack(side="left")
            status_label = tk.Label(bottom_row, text=str(spec["status"]), font=self._font(6), fg=APP_TEXT_SOFT, bg=APP_PANEL, wraplength=145, justify="left")
            status_label.pack(side="left", padx=(6, 0))
            percent_value = self._dashboard_metric_percent(str(spec["value"]), bool(spec["ok"]))
            percent_row = tk.Frame(inner_frame, bg=APP_PANEL)
            percent_row.pack(fill="x", padx=12, pady=(0, 10))
            track_widget, fill_widget = make_progress(percent_row, percent_value, APP_ACCENT if spec["ok"] else APP_WARNING)
            percent_label = tk.Label(percent_row, text=f"{percent_value}%", font=self._font(9, "bold", "Segoe UI Semibold"), fg=APP_ACCENT if spec["ok"] else APP_WARNING, bg=APP_PANEL)
            percent_label.pack(anchor="e", pady=(4, 0))
            self.dashboard_live_widgets[str(spec["key"])] = {
                "value_label": value_label,
                "status_label": status_label,
                "percent_label": percent_label,
                "track_widget": track_widget,
                "fill_widget": fill_widget,
            }

        alert_card = make_panel(metrics, bg="#241315", border="#533038", radius=20)
        alert_card.grid(row=0, column=4, sticky="nsew", padx=(8, 0))
        alert_body = alert_card.content  # type: ignore[attr-defined]
        alert_strip = tk.Frame(alert_body, bg=APP_DANGER if problem_count else APP_ACCENT, height=4)
        alert_strip.pack(fill="x", padx=12, pady=(10, 0))
        alert_inner = tk.Frame(
            alert_body,
            bg="#241315",
            bd=0,
            highlightthickness=1,
            highlightbackground="#4f2b30",
        )
        alert_inner.pack(fill="both", expand=True, padx=12, pady=(8, 10))
        alert_header = tk.Frame(alert_inner, bg="#241315")
        alert_header.pack(fill="x", padx=12, pady=(10, 6))
        icon_or_text(alert_header, "warning", "!", bg="#241315", fg=APP_DANGER)
        tk.Label(alert_header, text="Проверка на сигурност", font=self._font(7, "bold", "Segoe UI Semibold"), fg="#f3b1b1", bg="#241315").pack(side="left")
        alert_title_label = tk.Label(alert_inner, text="Внимание" if problem_count else "Няма проблем", font=self._font(11, "bold", "Segoe UI Semibold"), fg=APP_DANGER if problem_count else APP_ACCENT, bg="#241315")
        alert_title_label.pack(anchor="w", padx=12)
        alert_detail_label = tk.Label(
            alert_inner,
            text=f"{problem_count} проблем(а) открити" if problem_count else "Системата изглежда стабилна",
            font=self._font(6),
            fg="#ffd4d4" if problem_count else APP_TEXT_SOFT,
            bg="#241315",
        )
        alert_detail_label.pack(anchor="w", padx=12, pady=(8, 0))
        self.dashboard_live_widgets["security_alert"] = {
            "strip_widget": alert_strip,
            "title_label": alert_title_label,
            "detail_label": alert_detail_label,
        }

        lower = tk.Frame(outer, bg=APP_BG)
        lower.pack(fill="both", expand=True)
        lower.columnconfigure(0, weight=11, uniform="lower")
        lower.columnconfigure(1, weight=14, uniform="lower")
        lower.columnconfigure(2, weight=12, uniform="lower")

        info_panel = make_panel(lower, radius=20)
        info_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        info_panel_body = info_panel.content  # type: ignore[attr-defined]
        info_strip = tk.Frame(info_panel_body, bg=APP_ACCENT, height=4)
        info_strip.pack(fill="x", padx=12, pady=(10, 0))
        info_inner = tk.Frame(
            info_panel_body,
            bg=APP_PANEL,
            bd=0,
            highlightthickness=1,
            highlightbackground="#20352d",
        )
        info_inner.pack(fill="both", expand=True, padx=12, pady=(8, 8))
        info_header = tk.Frame(info_inner, bg=APP_PANEL)
        info_header.pack(fill="x", padx=12, pady=(10, 6))
        icon_or_text(info_header, "monitor_small", "▣", bg=APP_PANEL, fg=APP_ACCENT)
        tk.Label(info_header, text="Състояние на системата", font=self._font(12, "bold", "Segoe UI Semibold"), fg=APP_TEXT, bg=APP_PANEL).pack(side="left")
        info_scroll_host = tk.Frame(info_inner, bg=APP_PANEL)
        info_scroll_host.pack(fill="both", expand=True, padx=0, pady=(0, 0))
        info_scroll_host.columnconfigure(0, weight=1)
        info_scroll_host.rowconfigure(0, weight=1)
        info_canvas = tk.Canvas(
            info_scroll_host,
            bg=APP_PANEL,
            highlightthickness=0,
            bd=0,
            relief="flat",
            height=self._scale_px(360),
        )
        info_canvas.grid(row=0, column=0, sticky="nsew")
        info_viewport = tk.Frame(info_canvas, bg=APP_PANEL)
        self._populate_dashboard_info_rows(info_viewport, 235)
        self._bind_dashboard_info_mousewheel(info_viewport, info_canvas)
        info_window = info_canvas.create_window((0, 0), window=info_viewport, anchor="nw")

        # Помощна функция за refresh info scroll.
        def refresh_info_scroll(_: object | None = None) -> None:
            if not info_canvas.winfo_exists():
                return
            info_canvas.update_idletasks()
            content_height = max(info_viewport.winfo_reqheight(), 1)
            canvas_width = max(info_canvas.winfo_width(), 1)
            info_canvas.itemconfigure(info_window, width=canvas_width)
            info_canvas.configure(scrollregion=(0, 0, canvas_width, content_height))
            self._stop_dashboard_info_scroll()
            self.dashboard_info_scroll_job = self.root.after(
                1200,
                lambda: self._start_dashboard_info_scroll(info_canvas),
            )

        info_viewport.bind("<Configure>", refresh_info_scroll)
        info_canvas.bind("<Configure>", refresh_info_scroll)
        self._bind_dashboard_info_mousewheel(info_canvas, info_canvas)
        self.dashboard_live_widgets["system_info_scroll"] = {
            "canvas": info_canvas,
            "frame": info_viewport,
            "refresh_callback": refresh_info_scroll,
        }
        tk.Button(
            info_panel_body,
            text="Подробен системен отчет",
            command=lambda: self._handle_driver_backup_action("driver_pc_report"),
            font=self._font(9, "bold", "Segoe UI Semibold"),
            bg=APP_PANEL_ALT,
            fg=APP_TEXT,
            activebackground=APP_BORDER_STRONG,
            activeforeground="#ffffff",
            bd=0,
            padx=16,
            pady=12,
            cursor="hand2",
        ).pack(fill="x", padx=12, pady=(12, 12))

        installer_panel = make_panel(lower, radius=20)
        installer_panel.grid(row=0, column=1, sticky="nsew", padx=8)
        installer_panel_body = installer_panel.content  # type: ignore[attr-defined]
        installer_accent_strip = tk.Frame(installer_panel_body, bg=APP_ACCENT, height=4)
        installer_accent_strip.pack(fill="x", padx=12, pady=(10, 0))
        installer_inner = tk.Frame(
            installer_panel_body,
            bg=APP_PANEL,
            bd=0,
            highlightthickness=1,
            highlightbackground="#20352d",
        )
        installer_inner.pack(fill="both", expand=True, padx=12, pady=(8, 10))
        installer_header = tk.Frame(installer_inner, bg=APP_PANEL)
        installer_header.pack(fill="x", padx=12, pady=(10, 6))
        icon_or_text(installer_header, "robot_small", "▣", bg=APP_PANEL, fg=APP_ACCENT)
        title_box = tk.Frame(installer_header, bg=APP_PANEL)
        title_box.pack(side="left", fill="x", expand=True)
        tk.Label(title_box, text="Автоматичен инсталатор", font=self._font(12, "bold", "Segoe UI Semibold"), fg=APP_TEXT, bg=APP_PANEL).pack(anchor="w")
        tk.Label(title_box, text="Изберете какво да инсталирате", font=self._font(8), fg=APP_TEXT_SOFT, bg=APP_PANEL).pack(anchor="w", pady=(3, 0))
        tk.Button(
            installer_header,
            text="Управление на задачи",
            command=lambda: self.render_menu("auto_installer"),
            font=self._font(8, "bold", "Segoe UI Semibold"),
            bg=APP_ACCENT_SOFT,
            fg="#f2fff8",
            activebackground="#27a67a",
            activeforeground="#ffffff",
            bd=0,
            padx=14,
            pady=8,
            cursor="hand2",
        ).pack(side="right")
        list_holder = tk.Frame(installer_inner, bg=APP_PANEL_ALT)
        list_holder.pack(fill="both", expand=True, padx=12, pady=8)
        list_holder.columnconfigure(0, weight=1)
        list_holder.rowconfigure(0, weight=1)
        installer_canvas = tk.Canvas(
            list_holder,
            bg=APP_PANEL_ALT,
            highlightthickness=0,
            bd=0,
            relief="flat",
        )
        installer_canvas.grid(row=0, column=0, sticky="nsew")
        installer_viewport = tk.Frame(installer_canvas, bg=APP_PANEL_ALT)
        installer_window = installer_canvas.create_window((0, 0), window=installer_viewport, anchor="nw")

        # Помощна функция за refresh installer scroll.
        def refresh_installer_scroll(_: object | None = None) -> None:
            if not installer_canvas.winfo_exists():
                return
            installer_canvas.update_idletasks()
            content_height = max(installer_viewport.winfo_reqheight(), 1)
            canvas_width = max(installer_canvas.winfo_width(), 1)
            installer_canvas.itemconfigure(installer_window, width=canvas_width)
            installer_canvas.configure(scrollregion=(0, 0, canvas_width, content_height))

        installer_viewport.bind("<Configure>", refresh_installer_scroll)
        installer_canvas.bind("<Configure>", refresh_installer_scroll)
        self._bind_dashboard_canvas_mousewheel(installer_canvas, installer_canvas)
        preview_tasks = self._auto_install_tasks()
        self._ensure_auto_install_vars(preview_tasks)
        dashboard_count_var = tk.StringVar(value="Избрани задачи: 0")

        # Обновява dashboard install count след промяна в състоянието.
        def update_dashboard_install_count(*_: object) -> None:
            selected_count = sum(
                1
                for task in preview_tasks
                if self.auto_install_vars.get(task["id"]) and self.auto_install_vars[task["id"]].get()
            )
            dashboard_count_var.set(f"Избрани задачи: {selected_count}")

        # Задава dashboard install selection според избраното действие.
        def set_dashboard_install_selection(value: bool) -> None:
            for task in preview_tasks:
                task_id = task["id"]
                installed_now, _detail = self._dashboard_task_install_state(task)
                if installed_now:
                    self.auto_install_vars[task_id].set(False)
                    continue
                self.auto_install_vars[task_id].set(value)
            update_dashboard_install_count()

        category = ""
        visible_counter = 0
        for task in preview_tasks:
            if task["category"] != category:
                category = task["category"]
                tk.Label(installer_viewport, text=category, font=self._font(9, "bold", "Segoe UI Semibold"), fg=APP_ACCENT, bg=APP_PANEL_ALT).pack(anchor="w", padx=12, pady=(10, 4))
            installed_now, _ = self._dashboard_task_install_state(task)
            task_id = task["id"]
            if installed_now:
                self.auto_install_vars[task_id].set(False)
            visible_counter += 1
            row = tk.Frame(installer_viewport, bg=APP_PANEL_ALT)
            row.pack(fill="x", padx=12, pady=3)
            check_button = tk.Checkbutton(
                row,
                variable=self.auto_install_vars[task_id],
                text=f"{visible_counter}.  {task['label']}",
                font=self._font(8, "bold", "Segoe UI Semibold"),
                bg=APP_PANEL_ALT,
                activebackground=APP_PANEL_ALT,
                selectcolor="#174327",
                fg=APP_TEXT,
                activeforeground="#ffffff",
                anchor="w",
                bd=0,
                highlightthickness=0,
                padx=0,
                pady=0,
                command=update_dashboard_install_count,
            )
            if installed_now:
                check_button.config(state="disabled", fg=APP_TEXT_MUTED, disabledforeground=APP_TEXT_MUTED)
            check_button.pack(side="left", fill="x", expand=True, padx=(0, 10))
            pill_bg = "#133524" if installed_now else "#413412"
            pill_fg = APP_ACCENT if installed_now else "#f3bb4c"
            tk.Label(row, text="Наличен" if installed_now else "Частично", font=self._font(8, "bold", "Segoe UI Semibold"), fg=pill_fg, bg=pill_bg, padx=10, pady=3).pack(side="right")
            tk.Label(row, text="ⓘ", font=self._font(8), fg=APP_TEXT_MUTED, bg=APP_PANEL_ALT).pack(side="right", padx=(0, 8))
        self._bind_dashboard_canvas_mousewheel(installer_viewport, installer_canvas)
        update_dashboard_install_count()
        footer_row = tk.Frame(installer_inner, bg=APP_PANEL)
        footer_row.pack(fill="x", padx=12, pady=(0, 10))
        tk.Label(footer_row, textvariable=dashboard_count_var, font=self._font(9), fg=APP_TEXT_SOFT, bg=APP_PANEL).pack(side="left")
        tk.Button(
            footer_row,
            text="Избери всичко",
            command=lambda: set_dashboard_install_selection(True),
            font=self._font(8, "bold", "Segoe UI Semibold"),
            bg=APP_PANEL_ALT,
            fg=APP_TEXT,
            activebackground=APP_BORDER_STRONG,
            activeforeground="#ffffff",
            bd=0,
            padx=12,
            pady=6,
            cursor="hand2",
        ).pack(side="right", padx=(8, 0))
        tk.Button(
            footer_row,
            text="Изчисти",
            command=lambda: set_dashboard_install_selection(False),
            font=self._font(8, "bold", "Segoe UI Semibold"),
            bg=APP_PANEL_ALT,
            fg=APP_TEXT,
            activebackground=APP_BORDER_STRONG,
            activeforeground="#ffffff",
            bd=0,
            padx=12,
            pady=6,
            cursor="hand2",
        ).pack(side="right")
        tk.Button(
            installer_inner,
            text="▶  Стартирай инсталацията",
            command=self._start_auto_installer,
            font=self._font(11, "bold", "Segoe UI Semibold"),
            bg=APP_ACCENT_SOFT,
            fg="#f2fff8",
            activebackground="#27a67a",
            activeforeground="#ffffff",
            bd=0,
            padx=14,
            pady=13,
            cursor="hand2",
        ).pack(fill="x", padx=12, pady=(0, 12))

        right_column = tk.Frame(lower, bg=APP_BG)
        right_column.grid(row=0, column=2, sticky="nsew", padx=(8, 0))
        right_column.rowconfigure(0, weight=3)
        right_column.rowconfigure(1, weight=2)
        component_panel = make_panel(right_column, radius=20)
        component_panel.grid(row=0, column=0, sticky="nsew", pady=(0, 8))
        component_panel_body = component_panel.content  # type: ignore[attr-defined]
        component_strip = tk.Frame(component_panel_body, bg=APP_ACCENT, height=4)
        component_strip.pack(fill="x", padx=12, pady=(10, 0))
        component_inner = tk.Frame(
            component_panel_body,
            bg=APP_PANEL,
            bd=0,
            highlightthickness=1,
            highlightbackground="#20352d",
        )
        component_inner.pack(fill="both", expand=True, padx=12, pady=(8, 10))
        component_header = tk.Frame(component_inner, bg=APP_PANEL)
        component_header.pack(fill="x", padx=12, pady=(10, 6))
        icon_or_text(component_header, "shield_small", "▣", bg=APP_PANEL, fg=APP_ACCENT)
        tk.Label(component_header, text="Състояние на компонентите", font=self._font(12, "bold", "Segoe UI Semibold"), fg=APP_TEXT, bg=APP_PANEL).pack(side="left")
        component_holder = tk.Frame(component_inner, bg=APP_PANEL_ALT)
        component_holder.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        component_holder.columnconfigure(0, weight=1)
        component_holder.rowconfigure(0, weight=1)
        component_canvas = tk.Canvas(
            component_holder,
            bg=APP_PANEL_ALT,
            highlightthickness=0,
            bd=0,
            relief="flat",
            height=self._scale_px(250),
        )
        component_canvas.grid(row=0, column=0, sticky="nsew")
        component_viewport = tk.Frame(component_canvas, bg=APP_PANEL_ALT)
        component_window = component_canvas.create_window((0, 0), window=component_viewport, anchor="nw")

        # Помощна функция за refresh component scroll.
        def refresh_component_scroll(_: object | None = None) -> None:
            if not component_canvas.winfo_exists():
                return
            component_canvas.update_idletasks()
            content_height = max(component_viewport.winfo_reqheight(), 1)
            canvas_width = max(component_canvas.winfo_width(), 1)
            component_canvas.itemconfigure(component_window, width=canvas_width)
            component_canvas.configure(scrollregion=(0, 0, canvas_width, content_height))

        component_viewport.bind("<Configure>", refresh_component_scroll)
        component_canvas.bind("<Configure>", refresh_component_scroll)
        self._bind_dashboard_canvas_mousewheel(component_canvas, component_canvas)
        for label, value, ok in self._dashboard_component_rows_for_dashboard():
            row = tk.Frame(component_viewport, bg=APP_PANEL_ALT)
            row.pack(fill="x", padx=10, pady=3)
            inner = tk.Frame(row, bg=APP_PANEL_ALT)
            inner.pack(fill="x", padx=8, pady=6)
            tk.Label(inner, text=label, font=self._font(8), fg=APP_TEXT, bg=APP_PANEL_ALT, anchor="w", width=20, justify="left").pack(side="left")
            status_group = tk.Frame(inner, bg=APP_PANEL_ALT)
            status_group.pack(side="left", padx=(14, 0))
            tk.Label(status_group, text=value, font=self._font(8, "bold", "Segoe UI Semibold"), fg=APP_ACCENT if ok else APP_DANGER, bg=APP_PANEL_ALT, justify="left", anchor="w").pack(side="left")
            tk.Label(status_group, text="\u25c9" if ok else "!", font=self._font(10, "bold", "Segoe UI Semibold"), fg=APP_ACCENT if ok else APP_DANGER, bg=APP_PANEL_ALT).pack(side="left", padx=(8, 0))
        self._bind_dashboard_canvas_mousewheel(component_viewport, component_canvas)
        quick_panel = make_panel(right_column, radius=20)
        quick_panel.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        quick_panel_body = quick_panel.content  # type: ignore[attr-defined]
        quick_header = tk.Frame(quick_panel_body, bg=APP_PANEL)
        quick_header.pack(fill="x", padx=16, pady=(14, 10))
        icon_or_text(quick_header, "actions_small", "▣", bg=APP_PANEL, fg=APP_ACCENT)
        tk.Label(quick_header, text="Бързи действия", font=self._font(12, "bold", "Segoe UI Semibold"), fg=APP_TEXT, bg=APP_PANEL).pack(side="left")
        actions_frame = tk.Frame(quick_panel_body, bg=APP_PANEL)
        actions_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        for idx in range(4):
            actions_frame.columnconfigure(idx, weight=1, uniform="quick")
        for idx, action in enumerate(self._dashboard_quick_actions()):
            card = tk.Frame(actions_frame, bg=APP_PANEL_ALT, bd=0, highlightthickness=1, highlightbackground=APP_BORDER)
            card.grid(row=0, column=idx, sticky="nsew", padx=(0 if idx == 0 else 6, 0))
            if "menu" in action:
                command = lambda menu_key=action["menu"]: self.render_menu(menu_key)
            elif action["action_id"] == "open_console":
                command = lambda: messagebox.showinfo("Конзола", "Тази бърза конзола ще бъде вързана в следващата стъпка.", parent=self.root)
            else:
                command = lambda item={"kind": "action", "action_id": action["action_id"], "label": action["label"]}: self._handle_action(item)
            quick_icon = self.dashboard_icons.get(action["icon"])
            tk.Button(
                card,
                text=action["label"],
                command=command,
                font=self._font(9, "bold", "Segoe UI Semibold"),
                bg=APP_PANEL_ALT,
                fg=APP_TEXT,
                activebackground=APP_BORDER_STRONG,
                activeforeground="#ffffff",
                bd=0,
                padx=10,
                pady=12,
                cursor="hand2",
                justify="center",
                wraplength=120,
                image=quick_icon,
                compound="top",
            ).pack(fill="both", expand=True)

        footer = make_panel(outer, bg="#09110f", radius=18)
        footer.pack(fill="x", pady=(12, 0))
        footer_body = footer.content  # type: ignore[attr-defined]
        _, disk_ok = self._health_item_value("Disk C:")
        footer_items = [
            ("shield_small", "Сигурност", "Добра" if status_ok else "Риск", status_ok),
            ("refresh_small", "Производителност", "Отлична" if status_ok else "Провери", status_ok),
            ("drive_small", "Дисково състояние", "Добро" if disk_ok else "Внимание", disk_ok),
        ]
        for icon_name, title, value, ok in footer_items:
            segment = tk.Frame(footer_body, bg="#09110f")
            segment.pack(side="left", fill="x", expand=True, padx=16, pady=10)
            icon_or_text(segment, icon_name, "•", bg="#09110f", fg=APP_ACCENT)
            tk.Label(segment, text=f"{title}: ", font=self._font(10), fg=APP_TEXT_SOFT, bg="#09110f").pack(side="left")
            tk.Label(segment, text=value, font=self._font(10, "bold", "Segoe UI Semibold"), fg=APP_ACCENT if ok else APP_DANGER, bg="#09110f").pack(side="left", padx=(0, 10))
            dots = tk.Frame(segment, bg="#09110f")
            dots.pack(side="left")
            for dot_index in range(5):
                tk.Label(dots, text="●", font=self._font(8), fg=APP_ACCENT if ok or dot_index < 2 else "#2d3b36", bg="#09110f").pack(side="left", padx=1)

        self.page_label.config(text="Начален dashboard")
        self.prev_button.config(state="disabled")
        self.next_button.config(state="disabled")
        self.back_button.config(state="disabled")

    # Помощна функция за card columns.
    def _card_columns(self) -> int:
        current_width = self.root.winfo_width() or self.root.winfo_screenwidth()
        if self.current_menu == "main":
            if current_width < 1080:
                return 1
            if current_width < 1400:
                return 2
            return MAIN_CARD_COLUMNS
        if current_width < 1120:
            return 1
        return 2

    # Подготвя card според избраните настройки.
    def _build_card(self, parent: tk.Widget, item: dict[str, str]) -> tk.Frame:
        accent = self._card_accent(item)
        card_bg = APP_PANEL_SOFT if item["kind"] == "menu" else APP_PANEL_ALT
        border_color = APP_BORDER_STRONG if item["kind"] == "menu" else APP_BORDER
        has_remove_button = False
        if self._is_office_install_item(item):
            office_info = self._office_install_info(item["action_id"])
            has_remove_button = bool(office_info.installed and office_info.uninstall_string)
        card_height = self.scaled_menu_card_min_height.get(self.current_menu, self.scaled_card_min_height)
        if has_remove_button:
            card_height = max(card_height, self._scale_px(265))
        card = tk.Frame(
            parent,
            bg=card_bg,
            bd=0,
            highlightthickness=1,
            highlightbackground=border_color,
            height=card_height,
        )
        card.grid_propagate(False)

        top = tk.Frame(card, bg=card_bg)
        top.pack(fill="x", padx=self._scale_px(14), pady=(self._scale_px(12), self._scale_px(8)))

        icon_image = self._menu_icon_for_item(item, "card")
        if icon_image is not None:
            icon_box = tk.Frame(top, bg="#112925", bd=0, highlightthickness=1, highlightbackground=accent, width=self._scale_px(48), height=self._scale_px(48))
            icon_box.pack(side="left")
            icon_box.pack_propagate(False)
            tk.Label(icon_box, image=icon_image, bg="#112925").pack(expand=True)
        else:
            dot_size = self._scale_px(16)
            dot = tk.Canvas(top, width=dot_size, height=dot_size, bg=card_bg, highlightthickness=0)
            dot.create_oval(self._scale_px(2), self._scale_px(2), dot_size - self._scale_px(2), dot_size - self._scale_px(2), fill=accent, outline="")
            dot.pack(side="left")

        compact_text_menus = {"office_center", "nexus_admin", "office_install_center", "secret_install", "driver_backup", "language"}
        title_font = self._font(11 if self.current_menu in compact_text_menus else 12, "bold", "Segoe UI Semibold")
        title_wraplength = self.compact_card_title_wrap if self.current_menu in compact_text_menus else self.card_title_wrap
        title = tk.Label(
            top,
            text=item["label"],
            font=title_font,
            fg=APP_TEXT,
            bg=card_bg,
            anchor="w",
            wraplength=title_wraplength,
            justify="left",
        )
        title.pack(side="left", padx=(8, 0), fill="x", expand=True)

        description = self._item_description(item)
        desc_wraplength = self.compact_card_desc_wrap if self.current_menu in compact_text_menus else self.card_desc_wrap
        desc_label = tk.Label(
            card,
            text=description,
            font=self._font(9),
            fg=APP_TEXT_SOFT,
            bg=card_bg,
            wraplength=desc_wraplength,
            justify="left",
            anchor="nw",
        )
        desc_label.pack(fill="x", padx=self._scale_px(14), pady=(0, self._scale_px(8)))

        spacer = tk.Frame(card, bg=card_bg)
        spacer.pack(fill="both", expand=True)

        compact_action_menus = {"office_center", "office_install_center", "secret_install", "driver_backup", "language", "nexus_admin"}
        action_area_height = CARD_ACTION_DOUBLE_HEIGHT if has_remove_button else CARD_ACTION_HEIGHT
        action_area = tk.Frame(card, bg=card_bg, height=max(self.card_button_height_px + self._scale_px(12), self._scale_px(action_area_height)))
        action_pady = self._scale_px(8 if self.current_menu in compact_action_menus else 16)
        action_area.pack(fill="x", padx=self._scale_px(14), pady=(action_pady, action_pady), side="bottom")
        action_area.pack_propagate(False)

        action_text = self._button_text(item["kind"])
        button_bg, button_fg, button_active_bg = self._button_colors(item["kind"], accent)
        action_button = self._make_card_button(
            action_area,
            text=action_text,
            command=lambda selected=item: self.handle_item(selected),
            bg=button_bg,
            fg=button_fg,
            active_bg=button_active_bg,
            cursor="hand2" if item["kind"] != "info" else "arrow",
            state="disabled" if item["kind"] == "info" else "normal",
        )
        action_button.place(
            relx=0.5,
            y=0,
            anchor="n",
            width=self.card_button_width_px,
            height=self.card_button_height_px,
        )

        if self._is_office_install_item(item):
            if has_remove_button:
                remove_button = self._make_card_button(
                    action_area,
                    text="\u041f\u0440\u0435\u043c\u0430\u0445\u043d\u0438",
                    command=lambda selected=item: self._remove_office_installation(selected["action_id"]),
                    bg="#6b2730",
                    fg="#fff6f6",
                    active_bg="#8e3540",
                    cursor="hand2",
                )
                remove_button.place(
                    relx=0.5,
                    y=self.card_button_height_px + self.card_action_gap_px,
                    anchor="n",
                    width=self.card_button_width_px,
                    height=self.card_button_height_px,
                )

        return card

    # Помощна функция за is office install item.
    def _is_office_install_item(self, item: dict[str, str]) -> bool:
        action_id = item.get("action_id", "")
        return action_id.startswith("install_office_") and action_id.endswith("_offline")

    # Помощна функция за is office online item.
    def _is_office_online_item(self, item: dict[str, str]) -> bool:
        return item.get("action_id", "").startswith("online_")

    # Помощна функция за is office maintenance item.
    def _is_office_maintenance_item(self, item: dict[str, str]) -> bool:
        return item.get("action_id", "") in {
            "office_check_activation_status",
            "office_quick_repair",
            "office_force_uninstall_all",
        }

    # Помощна функция за is language item.
    def _is_language_item(self, item: dict[str, str]) -> bool:
        return item.get("action_id", "") in {
            "language_refresh",
            "toggle_bulgarian_bds",
            "toggle_bulgarian_phonetic",
            "toggle_bulgarian_traditional",
            "toggle_bulgarian_language_pack",
            "remove_bulgarian_language",
        }

    # Помощна функция за is driver backup item.
    def _is_driver_backup_item(self, item: dict[str, str]) -> bool:
        return item.get("action_id", "") in {
            "driver_backup_clean",
            "driver_backup_full",
            "driver_recovery_usb",
            "driver_pc_report",
            "driver_backup_advanced",
            "driver_restore_last",
        }

    # Помощна функция за is nexus admin item.
    def _is_nexus_admin_item(self, item: dict[str, str]) -> bool:
        return item.get("action_id", "") in {
            "nexus_list_users",
            "nexus_change_password",
            "nexus_create_user",
            "nexus_delete_user",
            "nexus_user_details",
            "nexus_toggle_admin",
        }

    # Помощна функция за office install info.
    def _office_install_info(self, action_id: str) -> object:
        if action_id not in self.office_inventory_cache:
            self.office_inventory_cache[action_id] = detect_installed_office(action_id)
        return self.office_inventory_cache[action_id]

    # Помощна функция за office online status.
    def _office_online_status(self, action_id: str) -> object:
        if action_id not in self.office_online_cache:
            self.office_online_cache[action_id] = check_online_package(action_id)
        return self.office_online_cache[action_id]

    # Помощна функция за office maintenance status.
    def _office_maintenance_status(self, action_id: str) -> object:
        if action_id not in self.office_maintenance_cache:
            self.office_maintenance_cache[action_id] = check_maintenance_action(action_id)
        return self.office_maintenance_cache[action_id]

    # Помощна функция за adobe reader status.
    def _adobe_reader_status(self) -> object:
        if self.adobe_reader_status_cache is None:
            self.adobe_reader_status_cache = check_adobe_reader_status(PROJECT_ROOT)
        return self.adobe_reader_status_cache

    # Помощна функция за language status.
    def _language_status(self) -> object:
        if self.language_status_cache is None:
            self.language_status_cache = get_language_status()
        return self.language_status_cache

    # Помощна функция за reset language status cache.
    def _reset_language_status_cache(self) -> None:
        self.language_status_cache = None

    # Помощна функция за last driver backup dir.
    def _last_driver_backup_dir(self) -> Path | None:
        last_backup = self.settings.get("last_driver_backup_dir", "")
        if last_backup:
            backup_path = Path(last_backup)
            if backup_path.exists():
                return backup_path
        return None

    # Помощна функция за nexus admin status.
    def _nexus_admin_status(self) -> object:
        if self.nexus_admin_status_cache is None:
            self.nexus_admin_status_cache = check_nexus_admin_status()
        return self.nexus_admin_status_cache

    # Събира command output от системата.
    def _collect_command_output(self, result: subprocess.CompletedProcess[str]) -> str:
        # Събира stdout и stderr в един удобен текст.
        return "\n".join(part.strip() for part in (result.stdout, result.stderr) if part and part.strip())

    # Помощна функция за append command output.
    def _append_command_output(self, output_text: str) -> None:
        # Добавя текст от команда в прозореца за прогрес, ако има такъв.
        if output_text.strip():
            self.root.after(0, lambda text=output_text.strip(): self._append_activation_log(text))

    # Помощна функция за is winget package installed.
    def _is_winget_package_installed(self, package_id: str) -> tuple[bool, str]:
        # Проверява дали даден winget пакет вече е инсталиран.
        winget_exe = find_winget_executable()
        if not winget_exe:
            return False, ""
        result = subprocess.run(
            [winget_exe, "list", "--id", package_id, "--source", "winget"],
            capture_output=True,
            text=True,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        output = self._collect_command_output(result)
        normalized = output.lower()
        if "no installed package found" in normalized or "no package found matching input criteria" in normalized:
            return False, output
        return package_id.lower() in normalized, output

    # Помощна функция за office online install state.
    def _office_online_install_state(self, action_id: str) -> tuple[bool, str, str]:
        # Търси в registry дали този Office пакет вече го има и връща и команда за махане.
        package = get_online_package(action_id)
        installed, details, uninstall_string = self._find_installed_registry_app(package.registry_patterns)
        return installed, details or package.label, uninstall_string

    # Помощна функция за office install architecture.
    def _office_install_architecture(self) -> str:
        # Избира 64-bit при 64-bit Windows, а 32-bit само ако системата е 32-bitова.
        return "64" if os.environ.get("ProgramFiles(x86)") else "32"

    # Изтегля office deployment tool от зададения адрес.
    def _download_office_deployment_tool(self, target_dir: Path) -> Path:
        # Дърпа последния ODT installer от официалната Microsoft страница.
        request = urllib.request.Request(
            ODT_CONFIRMATION_URL,
            headers={"User-Agent": "Mozilla/5.0 WinSysGuardianAdvanced"},
        )
        with urllib.request.urlopen(request, timeout=30) as response:
            confirmation_html = response.read().decode("utf-8", errors="replace")

        match = re.search(
            r"https://download\.microsoft\.com/download/[^\"'<>\\s]+officedeploymenttool_[^\"'<>\\s]+\.exe",
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

    # Извлича office deployment tool от подадения текст или архив.
    def _extract_office_deployment_tool(self, odt_exe: Path, target_dir: Path) -> Path:
        # Разпакрира ODT и връща setup.exe, което реално пуска online инсталацията.
        command = [str(odt_exe), "/quiet", f"/extract:{target_dir}"]
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        setup_path = target_dir / "setup.exe"
        if setup_path.exists():
            return setup_path
        output = self._collect_command_output(result)
        raise RuntimeError(output or "Office Deployment Tool не се разпакрира правилно.")

    # Помощна функция за write office online config.
    def _write_office_online_config(self, config_path: Path, package: object) -> None:
        # Създава временния XML файл за точната online Office инсталация.
        config_text = f"""<Configuration>
  <Add OfficeClientEdition="{self._office_install_architecture()}" Channel="{package.channel}">
    <Product ID="{package.product_id}">
      <Language ID="MatchOS" Fallback="en-us" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"""
        config_path.write_text(config_text, encoding="utf-8")

    # Стартира office online install core и връща резултата.
    def _run_office_online_install_core(self, action_id: str, remove_existing: bool = False) -> str:
        # Това е общата логика за online Office, за да работи и при един бутон, и при автоматичния installer.
        package = get_online_package(action_id)
        status = check_online_package(action_id)
        if not status.available:
            raise RuntimeError(status.message)

        installed_now, installed_details, uninstall_string = self._office_online_install_state(action_id)
        output_lines: list[str] = []

        with tempfile.TemporaryDirectory(prefix="wga-office-online-") as temp_dir_name:
            temp_dir = Path(temp_dir_name)
            odt_dir = temp_dir / "odt"
            odt_dir.mkdir(parents=True, exist_ok=True)

            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    15,
                    f"Подготовка на online инсталацията за {package.label}...",
                    "Търси се последният Office Deployment Tool от Microsoft.",
                ),
            )
            odt_exe = self._download_office_deployment_tool(odt_dir)
            output_lines.append(f"Office Deployment Tool: {odt_exe.name}")

            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    35,
                    "Изтегленият Office Deployment Tool се разархивира...",
                    str(odt_exe),
                ),
            )
            setup_exe = self._extract_office_deployment_tool(odt_exe, odt_dir)

            if remove_existing and installed_now and uninstall_string:
                self.root.after(
                    0,
                    lambda: self._update_activation_progress(
                        50,
                        f"Премахване на стара версия за {package.label}...",
                        installed_details,
                    ),
                )
                removal_text = self._run_office_uninstall_command(action_id, installed_details, uninstall_string)
                if removal_text:
                    output_lines.append(removal_text)

            config_path = temp_dir / "configuration.xml"
            self._write_office_online_config(config_path, package)
            command = [str(setup_exe), "/configure", str(config_path)]

            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    70,
                    f"Стартиране на online Office инсталация за {package.label}...",
                    f"Product ID: {package.product_id}\nChannel: {package.channel}",
                ),
            )
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
                cwd=str(odt_dir),
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            output = self._collect_command_output(result)
            if output:
                output_lines.append(output)
                self.root.after(0, lambda text=output: self._append_activation_log(text))
            if result.returncode != 0:
                raise RuntimeError("\n\n".join(output_lines) or f"{package.label} върна код {result.returncode}.")

        self.office_online_cache.pop(action_id, None)
        self.office_inventory_cache.pop(action_id, None)
        return "\n\n".join(output_lines) or f"{package.label} стартира успешно чрез Office Deployment Tool."

    # Стартира office uninstall command и връща резултата.
    def _run_office_uninstall_command(self, action_id: str, display_name: str, uninstall_string: str) -> str:
        # Изпълнява реалната деинсталация на намерен Office пакет.
        result = subprocess.run(
            ["cmd", "/c", uninstall_string],
            capture_output=True,
            text=True,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        output = self._collect_command_output(result)
        self._append_command_output(output)
        if result.returncode != 0:
            raise RuntimeError(output or f"Uninstall command returned code {result.returncode}.")
        self.office_inventory_cache.pop(action_id, None)
        return output or f"{display_name} removal finished."

    # Стартира winget uninstall command и връща резултата.
    def _run_winget_uninstall_command(self, winget_exe: str, package_id: str, label: str) -> str:
        # Премахва winget пакет преди нова инсталация.
        result = subprocess.run(
            [winget_exe, "uninstall", "--id", package_id, "--silent", "--accept-source-agreements"],
            capture_output=True,
            text=True,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        output = self._collect_command_output(result)
        self._append_command_output(output)
        normalized = output.lower()
        if result.returncode != 0 and "no installed package found" not in normalized and "no package found matching input criteria" not in normalized:
            raise RuntimeError(output or f"{label} uninstall returned code {result.returncode}.")
        return output or f"{label} removal finished."

    # Помощна функция за item description.
    def _item_description(self, item: dict[str, str]) -> str:
        description = item.get("description", self._kind_description(item["kind"]))
        if self._is_office_install_item(item):
            office_info = self._office_install_info(item["action_id"])
            installer = get_office_offline_installer(item["action_id"])
            status_line = (
                f"\n\n✓ Инсталирано: {office_info.display_name}"
                if office_info.installed
                else "\n\n✗ Не е открита инсталирана версия."
            )
            installers_line = (
                f"\n✓ Папка с инсталатори: {installer.installers_root}"
                if installer.installers_root.exists()
                else f"\n✗ Липсва папка с инсталатори: {installer.installers_root}"
            )
            return f"{description}{status_line}{installers_line}"

        if self._is_office_online_item(item):
            online_status = self._office_online_status(item["action_id"])
            marker = "✓" if online_status.available else "✗"
            base_description = item.get(
                "description",
                "Проверява дали пакетът е наличен онлайн чрез winget и дали може да се инсталира.",
            )
            return f"{base_description}\n\n{marker} {online_status.message}"

        if self._is_office_maintenance_item(item):
            maintenance_status = self._office_maintenance_status(item["action_id"])
            marker = "✓" if maintenance_status.available else "✗"
            return f"{description}\n\n{marker} {maintenance_status.message}"

        if item.get("action_id") == "install_adobe_reader":
            status = self._adobe_reader_status()
            latest = getattr(status, "latest_version", "") or "неизвестна"
            local_path = getattr(status, "local_installer", None)
            local_line = "локален файл: OK" if local_path else "локален файл: липсва"
            return (
                f"Проверява Adobe Reader през winget.\n"
                f"Версия: {latest} | {local_line}"
            )

        if self._is_language_item(item):
            try:
                language_status = self._language_status()
            except Exception as exc:
                return "Статусът не може да се провери."

            action_id = item.get("action_id", "")
            if action_id == "language_refresh":
                ready_count = sum(
                    [
                        language_status.has_bulgarian,
                        language_status.has_language_pack,
                        language_status.has_bds,
                        language_status.has_phonetic,
                        language_status.has_traditional,
                    ]
                )
                return f"Налични: {ready_count}/5"
            if action_id == "toggle_bulgarian_bds":
                state = "налична" if language_status.has_bds else "липсва"
                return f"Статус: {state}"
            if action_id == "toggle_bulgarian_phonetic":
                state = "налична" if language_status.has_phonetic else "липсва"
                return f"Статус: {state}"
            if action_id == "toggle_bulgarian_traditional":
                state = "налична" if language_status.has_traditional else "липсва"
                return f"Статус: {state}"
            if action_id == "toggle_bulgarian_language_pack":
                state = "наличен" if language_status.has_language_pack else "липсва"
                return f"Статус: {state}"
            if action_id == "remove_bulgarian_language":
                state = "bg-BG е наличен" if language_status.has_bulgarian else "bg-BG не е намерен"
                return state

        if self._is_driver_backup_item(item):
            action_id = item.get("action_id", "")
            last_backup_dir = self._last_driver_backup_dir()
            if action_id in {"driver_backup_clean", "driver_backup_full"}:
                marker = "✓" if last_backup_dir else "✗"
                suffix = f"Last backup: {last_backup_dir}" if last_backup_dir else "No previous driver backup recorded yet."
                return f"{description}\n\n{marker} {suffix}"
            if action_id == "driver_recovery_usb":
                usb_drives = detect_removable_drives()
                backup_ok = last_backup_dir is not None
                return (
                    f"{description}\n\n"
                    f"{'✓' if backup_ok else '✗'} Last backup {'found' if backup_ok else 'missing'}\n"
                    f"{'✓' if usb_drives else '✗'} Removable USB {'detected' if usb_drives else 'not detected'}"
                )
            if action_id == "driver_pc_report":
                last_report = self.settings.get("last_pc_report_path", "")
                marker = "✓" if last_report and Path(last_report).exists() else "✗"
                suffix = f"Last report: {last_report}" if marker == "✓" else "No PC report generated yet."
                return f"{description}\n\n{marker} {suffix}"
            if action_id == "driver_backup_advanced":
                usb_drives = detect_removable_drives()
                onedrive_dir = onedrive_path()
                return (
                    f"{description}\n\n"
                    f"✓ Desktop always available\n"
                    f"{'✓' if usb_drives else '✗'} USB destination\n"
                    f"{'✓' if onedrive_dir else '✗'} OneDrive destination\n"
                    f"✓ NAS path can be entered manually"
                )
            if action_id == "driver_restore_last":
                marker = "✓" if last_backup_dir else "✗"
                suffix = f"Ready to restore from: {last_backup_dir}" if last_backup_dir else "No backup folder is saved yet."
                return f"{description}\n\n{marker} {suffix}"

        if self._is_nexus_admin_item(item):
            nexus_status = self._nexus_admin_status()
            marker = "✓" if nexus_status.available else "✗"
            return f"{marker} {nexus_status.message}"

        return description

    # Помощна функция за card accent.
    def _card_accent(self, item: dict[str, str]) -> str:
        if item.get("accent"):
            return item["accent"]
        accent_map = {
            "menu": "#2ea8ff",
            "action": "#39c25a",
            "exit": "#d94d4d",
            "info": "#8c9aa3",
        }
        return accent_map.get(item["kind"], "#39c25a")

    # Помощна функция за button colors.
    def _button_colors(self, kind: str, accent: str) -> tuple[str, str, str]:
        if kind == "menu":
            return (APP_ACCENT_BLUE, "#f4fbff", "#46a4ff")
        if kind == "exit":
            return ("#6b2730", "#fff6f6", "#8e3540")
        if kind == "info":
            return ("#263632", "#d8e2db", "#263632")
        return (APP_ACCENT_SOFT, "#f5fff7", "#27a67a")

    # Помощна функция за kind description.
    def _kind_description(self, kind: str) -> str:
        descriptions = {
            "menu": "Open this module and view its available tools.",
            "action": "Prepared action placeholder. We can connect it to a real script next.",
            "exit": "Close the current application session.",
            "info": "Information card for system status or guidance.",
        }
        return descriptions.get(kind, "Module item")

    # Помощна функция за button text.
    def _button_text(self, kind: str) -> str:
        labels = {
            "menu": "Enter Menu",
            "action": "Run",
            "exit": "Exit",
            "info": "Info",
        }
        return labels.get(kind, "Open")

    # Обработва събитието handle item.
    def handle_item(self, item: dict[str, str]) -> None:
        kind = item["kind"]
        if kind == "menu":
            target = item["target"]
            if target == "main":
                self.go_dashboard()
                return
            if target != self.current_menu:
                self.history.append(self.current_menu)
            self.render_menu(target)
        elif kind == "action":
            self._handle_action(item)
        elif kind == "exit":
            self.root.destroy()

    # Помощна функция за authorize windows11 menu.
    def _authorize_windows11_menu(self) -> bool:
        # Менюто за Windows 11 вече е без парола и се отваря директно.
        return True

    # Обработва събитието handle action.
    def _handle_action(self, item: dict[str, str]) -> None:
        action_id = item.get("action_id", "")
        if action_id == "add_desktop_icons":
            self._add_desktop_icons()
            return
        if action_id == "open_program_selector":
            self.render_menu("auto_installer")
            return
        if action_id == "save_windows11_key":
            self._save_windows11_key()
            return
        if action_id == "save_windows10_key":
            self._save_windows10_key()
            return
        if action_id == "activate_windows10":
            self._activate_windows10()
            return
        if action_id == "activate_windows11":
            self._activate_windows11()
            return
        if action_id == "show_windows10_key":
            self._show_windows10_key()
            return
        if action_id == "clear_windows10_key":
            self._clear_windows10_key()
            return
        if action_id == "show_windows11_key":
            self._show_windows11_key()
            return
        if action_id == "clear_windows11_key":
            self._clear_windows11_key()
            return
        if action_id == "save_office_key":
            self._save_office_key()
            return
        if action_id == "show_office_key":
            self._show_office_key()
            return
        if action_id == "clear_office_key":
            self._clear_office_key()
            return
        if action_id in {"office_2016_activation", "office_2019_activation", "office_2021_activation"}:
            self._activate_office_version(action_id)
            return
        if action_id in {"reset_onedrive_1", "reset_onedrive_2", "reset_onedrive_3"}:
            self._reset_onedrive(action_id)
            return
        if action_id in {
            "language_refresh",
            "toggle_bulgarian_bds",
            "toggle_bulgarian_phonetic",
            "toggle_bulgarian_traditional",
            "toggle_bulgarian_language_pack",
            "remove_bulgarian_language",
        }:
            self._handle_language_action(action_id, item["label"])
            return
        if action_id.startswith("install_office_") and action_id.endswith("_offline"):
            self._install_office_offline(action_id)
            return
        if action_id in {
            "driver_backup_clean",
            "driver_backup_full",
            "driver_recovery_usb",
            "driver_pc_report",
            "driver_backup_advanced",
            "driver_restore_last",
        }:
            self._handle_driver_backup_action(action_id)
            return
        if action_id == "install_adobe_reader":
            self._install_adobe_reader()
            return
        if action_id == "install_ninite":
            self._install_local_installer("install_ninite")
            return
        if action_id in {"install_visual_studio_setup", "install_vscode_arm64"}:
            self._install_local_installer(action_id)
            return
        if action_id == "office_check_activation_status":
            self._check_office_activation_status()
            return
        if action_id in {
            "nexus_list_users",
            "nexus_change_password",
            "nexus_create_user",
            "nexus_delete_user",
            "nexus_user_details",
            "nexus_toggle_admin",
        }:
            self._handle_nexus_admin_action(action_id)
            return
        if action_id == "office_quick_repair":
            self._quick_repair_office()
            return
        if action_id == "office_force_uninstall_all":
            self._force_uninstall_all_office()
            return
        if action_id.startswith("online_"):
            self._install_office_online(action_id)
            return
        if action_id == "hidden_show_status":
            messagebox.showinfo(
                "Hidden Menu",
                "Скритото меню работи и приложението е готово за нови действия.",
                parent=self.root,
            )
            return
        if action_id == "hidden_load_agent_status":
            self._show_agent_status()
            return

        self.status_var.set(f"Selected action: {item['label']}. This is ready to connect to a real Python or PowerShell task.")

    # Помощна функция за shortcut launch parts.
    def _shortcut_launch_parts(self, menu_key: str | None = None) -> tuple[str, str, str]:
        # Подготвя как да стартира приложението от shortcut според това дали е build или проект.
        if getattr(sys, "frozen", False):
            target_path = str(Path(sys.executable).resolve())
            arguments = f'--menu {menu_key}' if menu_key else ""
            working_dir = str(PROJECT_ROOT)
            return target_path, arguments, working_dir

        script_path = str(Path(__file__).resolve())
        target_path = sys.executable
        base_args = [script_path]
        if menu_key:
            base_args.extend(["--menu", menu_key])
        arguments = subprocess.list2cmdline(base_args)
        working_dir = str(Path(__file__).resolve().parent)
        return target_path, arguments, working_dir

    # Създава windows shortcut и връща резултата към приложението.
    def _create_windows_shortcut(self, shortcut_path: Path, target_path: str, arguments: str, working_dir: str) -> None:
        # Създава .lnk shortcut през PowerShell и WScript.Shell.
        icon_path = str(APP_ICON_FILE.resolve())
        escaped_shortcut = str(shortcut_path).replace("'", "''")
        escaped_target = target_path.replace("'", "''")
        escaped_arguments = arguments.replace("'", "''")
        escaped_working_dir = working_dir.replace("'", "''")
        escaped_icon = icon_path.replace("'", "''")
        script = (
            "$WshShell = New-Object -ComObject WScript.Shell; "
            f"$Shortcut = $WshShell.CreateShortcut('{escaped_shortcut}'); "
            f"$Shortcut.TargetPath = '{escaped_target}'; "
            f"$Shortcut.Arguments = '{escaped_arguments}'; "
            f"$Shortcut.WorkingDirectory = '{escaped_working_dir}'; "
            f"$Shortcut.IconLocation = '{escaped_icon}'; "
            "$Shortcut.Save()"
        )
        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
            check=True,
            capture_output=True,
            text=True,
        )

    # Помощна функция за add desktop icons.
    def _add_desktop_icons(self) -> None:
        # Пуска отделен прозорец с прогрес, докато Windows системните икони се включват.
        confirmed = messagebox.askyesno(
            "Add Desktop Icons",
            "Add the standard system icons This PC, Network, Control Panel and User Files to the desktop?",
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set("Adding desktop icons was canceled.")
            return

        self.status_var.set("Adding desktop system icons...")
        self._open_activation_window(
            title="Add Desktop Icons",
            heading="Системни икони на работния плот",
            intro="Приложението включва системните икони на Windows и опреснява работния плот.",
        )
        threading.Thread(target=self._run_add_desktop_icons, daemon=True).start()

    # Стартира add desktop icons и връща резултата.
    def _run_add_desktop_icons(self) -> None:
        # Работи във фонов режим, за да не блокира интерфейса по време на промяната.
        try:
            enabled_labels = enable_windows_desktop_icons(
                lambda value, status, details: self.root.after(
                    0,
                    lambda v=value, s=status, d=details: self._update_activation_progress(v, s, d),
                )
            )
        except OSError as exc:
            self.root.after(0, lambda: self._show_activation_result(False, str(exc), "Desktop Icons"))
            self.root.after(0, lambda: self.status_var.set("Неуспешно добавяне на системни икони на работния плот."))
            return

        summary = "Активирани икони:\n" + "\n".join(f"- {label}" for label in enabled_labels)
        self.root.after(0, lambda: self._show_activation_result(True, summary, "Desktop Icons"))
        self.root.after(0, lambda: self.status_var.set("Системните икони на работния плот бяха добавени успешно."))

    # Помощна функция за start auto installer.
    def _start_auto_installer(self) -> None:
        if self.auto_install_running:
            messagebox.showinfo("Автоматичен инсталатор", "Вече има стартирана автоматична инсталация.", parent=self.root)
            return

        tasks = []
        for task in self._auto_install_tasks():
            task_id = task["id"]
            if self.auto_install_vars.get(task_id) and self.auto_install_vars[task_id].get():
                prepared_task = dict(task)
                prepared_task["remove_first"] = bool(self.auto_remove_vars.get(task_id) and self.auto_remove_vars[task_id].get())
                tasks.append(prepared_task)
        if not tasks:
            messagebox.showinfo("Auto Installer", "Select at least one install task.", parent=self.root)
            return

        confirmed = messagebox.askyesno(
            "Auto Installer",
            f"{len(tasks)} task(s) will run one after another.\n\nStart now?",
            parent=self.root,
        )
        if not confirmed:
            return

        self._close_program_selector_window()
        self.auto_install_running = True
        self.status_var.set("Auto Installer is starting...")
        self._open_activation_window(
            title="Автоматичен инсталатор",
            heading="Автоматичен инсталатор",
            intro="Избраните задачи се изпълняват последователно. Не затваряй приложението, докато процесът работи.",
        )
        threading.Thread(target=self._run_auto_installer, args=(tasks,), daemon=True).start()

    # Стартира auto installer и връща резултата.
    def _run_auto_installer(self, tasks: list[dict[str, str]]) -> None:
        results: list[str] = []
        failures = 0
        total = len(tasks)
        for index, task in enumerate(tasks, start=1):
            base_progress = int((index - 1) * 90 / max(1, total))
            self.root.after(
                0,
                lambda task=task, index=index, total=total, base_progress=base_progress: self._update_activation_progress(
                    base_progress,
                    f"Задача {index}/{total}: {task['label']}",
                    f"Стартиране: {task['label']}",
                ),
            )
            try:
                detail = self._run_auto_install_task(task)
                results.append(f"✓ {task['label']}\n{detail}")
                self.root.after(0, lambda task=task: self._append_activation_log(f"Готово: {task['label']}"))
            except Exception as exc:
                failures += 1
                results.append(f"✗ {task['label']}\n{exc}")
                self.root.after(0, lambda task=task, exc=exc: self._append_activation_log(f"Проблем: {task['label']}\n{exc}"))

        success = failures == 0
        summary = "\n\n".join(results)
        subject = "Автоматичен инсталатор"

        # Помощна функция за finish.
        def finish() -> None:
            self.auto_install_running = False
            self._show_activation_result(success, summary, subject)
            self._refresh_resource_panel()
            if self.current_menu == "auto_installer":
                self._render_cards()

        self.root.after(0, finish)

    # Стартира auto language action и връща резултата.
    def _run_auto_language_action(self, action_id: str) -> str:
        status = get_language_status()
        already_ready = {
            "toggle_bulgarian_language_pack": status.has_language_pack,
            "toggle_bulgarian_bds": status.has_bds,
            "toggle_bulgarian_phonetic": status.has_phonetic,
            "toggle_bulgarian_traditional": status.has_traditional,
        }
        if already_ready.get(action_id, False):
            return "Компонентът вече е наличен. Пропуснато."

        title, command = build_language_action(action_id, status)
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command],
            capture_output=True,
            text=True,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        output = "\n".join(part.strip() for part in (result.stdout, result.stderr) if part and part.strip())
        if result.returncode != 0:
            raise RuntimeError(output or f"{title} върна код {result.returncode}.")
        self.language_status_cache = None
        return output or f"{title} завърши успешно."

    # Помощна функция за auto install tasks.
    def _auto_install_tasks(self) -> list[dict[str, str]]:
        if self.program_selector_tasks_cache:
            return [dict(task) for task in self.program_selector_tasks_cache]
        # Събира всички налични инсталации за прозореца с тикчета.
        tasks: list[dict[str, str]] = [dict(task) for task in PROGRAM_SELECTOR_LOCAL_TASKS]
        known_resource_ids = {
            task["resource_id"]
            for task in PROGRAM_SELECTOR_LOCAL_TASKS
            if task.get("resource_id")
        }

        for installer in OFFICE_OFFLINE_INSTALLERS.values():
            tasks.append(
                {
                    "id": installer.action_id,
                    "label": installer.label,
                    "category": "Office offline",
                    "description": f"Инсталация от Installers папката: {installer.folder}",
                    "type": "office_offline",
                }
            )

        tasks.append(
            {
                "id": "install_adobe_reader",
                "label": "Adobe Reader",
                "category": "Основен софтуер",
                "description": "Инсталира или обновява Adobe Reader през winget.",
                "type": "adobe",
            }
        )

        for package in OFFICE_ONLINE_PACKAGES.values():
            tasks.append(
                {
                    "id": package.action_id,
                    "label": package.label,
                    "category": "Office online",
                    "description": f"Online инсталация през winget: {package.winget_id}",
                    "type": "office_online",
                }
            )
            tasks[-1]["description"] = f"Online инсталация през ODT: {package.product_id or 'неподдържан пакет'}"

        for item in load_resource_manifest(PROJECT_ROOT):
            if item.resource_id in known_resource_ids:
                continue
            if item.resource_id == "adobe_reader":
                continue
            if item.resource_id.startswith("office_"):
                continue
            tasks.append(
                {
                    "id": f"resource_{item.resource_id}",
                    "label": item.name,
                    "category": item.category or "Допълнителни ресурси",
                    "description": f"Локален ресурс в Installers папката: {item.required_files[0]}",
                    "type": "resource_info",
                    "resource_id": item.resource_id,
                }
            )

        for action_id, label in (
            ("toggle_bulgarian_language_pack", "Български езиков пакет"),
            ("toggle_bulgarian_bds", "Българска БДС клавиатура"),
            ("toggle_bulgarian_phonetic", "Българска фонетична клавиатура"),
            ("toggle_bulgarian_traditional", "Традиционна фонетична клавиатура"),
        ):
            tasks.append(
                {
                    "id": action_id,
                    "label": label,
                    "category": "Език и клавиатури",
                    "description": "Добавя компонента само ако липсва.",
                    "type": "language",
                }
            )
        return tasks

    # Помощна функция за dashboard task install state.
    def _dashboard_task_install_state(self, task: dict[str, str]) -> tuple[bool, str]:
        cached = self.program_selector_status_cache.get(task["id"])
        if cached is not None:
            return cached
        self._refresh_program_selector_status_async()
        return False, "Проверява се..."

    # Помощна функция за refresh program selector status async.
    def _refresh_program_selector_status_async(self) -> None:
        if self.program_selector_scan_running:
            return
        self.program_selector_scan_running = True

        # Помощна функция за worker.
        def worker() -> None:
            try:
                tasks = self._auto_install_tasks()
                status_map: dict[str, tuple[bool, str]] = {}
                for task in tasks:
                    status_map[task["id"]] = self._safe_task_install_state(task)
            except Exception:
                tasks = self.program_selector_tasks_cache or []
                status_map = dict(self.program_selector_status_cache)
            try:
                self.root.after(0, lambda: self._apply_program_selector_status_cache(tasks, status_map))
            except RuntimeError:
                self.program_selector_scan_running = False

        threading.Thread(target=worker, daemon=True).start()

    # Помощна функция за apply program selector status cache.
    def _apply_program_selector_status_cache(
        self,
        tasks: list[dict[str, str]],
        status_map: dict[str, tuple[bool, str]],
    ) -> None:
        self.program_selector_scan_running = False
        if tasks:
            self.program_selector_tasks_cache = [dict(task) for task in tasks]
        self.program_selector_status_cache = dict(status_map)
        if self.current_menu == "main":
            self._render_cards()

    # Помощна функция за local task spec.
    def _local_task_spec(self, task_id: str) -> dict[str, str] | None:
        # Връща настройките за локален installer, ако задачата е такава.
        for task in PROGRAM_SELECTOR_LOCAL_TASKS:
            if task["id"] == task_id:
                return dict(task)
        return None

    # Помощна функция за manifest items by id.
    def _manifest_items_by_id(self) -> dict[str, object]:
        # Зарежда manifest елементите в удобен речник по ID.
        return {item.resource_id: item for item in load_resource_manifest(PROJECT_ROOT)}

    # Намира resource local file.
    def _find_resource_local_file(self, resource_id: str) -> Path | None:
        # Намира първия наличен локален файл за даден ресурс.
        item = self._manifest_items_by_id().get(resource_id)
        if not item:
            return None
        installers_root = self.resource_status.installers_root
        for relative_path in item.required_files:
            candidate = installers_root / relative_path
            if candidate.exists():
                return candidate
        return None

    # Намира installed registry app.
    def _find_installed_registry_app(self, patterns: tuple[str, ...]) -> tuple[bool, str, str]:
        # Търси програма в Windows registry и връща име, версия и команда за махане.
        if not patterns:
            return False, "", ""
        regex = "|".join("(" + pattern.replace("'", "''") + ")" for pattern in patterns)
        script = (
            "$paths=@("
            "'HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*',"
            "'HKLM:\\Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*',"
            "'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*'"
            ");"
            "$item=Get-ItemProperty -Path $paths -ErrorAction SilentlyContinue | "
            f"Where-Object {{ $_.DisplayName -and $_.DisplayName -match '{regex}' }} | "
            "Select-Object -First 1 DisplayName,DisplayVersion,UninstallString;"
            "if($item){$item|ConvertTo-Json -Compress}"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
            capture_output=True,
            text=True,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        output = self._collect_command_output(result).strip()
        if not output:
            return False, "", ""
        try:
            data = json.loads(output)
        except json.JSONDecodeError:
            return False, "", ""
        name = str(data.get("DisplayName") or "").strip()
        version = str(data.get("DisplayVersion") or "").strip()
        uninstall_string = str(data.get("UninstallString") or "").strip()
        if not name:
            return False, "", ""
        return True, f"{name} {version}".strip(), uninstall_string

    # Помощна функция за adobe install state.
    def _adobe_install_state(self) -> tuple[bool, str, str]:
        # Събира на едно място проверката за Adobe, за да работи и без winget запис.
        installed, detail, uninstall_string = self._find_installed_registry_app(
            ("Adobe Acrobat.*Reader", "Adobe Reader", "Acrobat Reader"),
        )
        if installed:
            return True, detail, uninstall_string
        status = self._adobe_reader_status()
        installed_version = getattr(status, "installed_version", "") or ""
        return bool(installed_version), installed_version or "Adobe Reader", ""

    # Стартира uninstall string command и връща резултата.
    def _run_uninstall_string_command(self, display_name: str, uninstall_string: str) -> str:
        # Стартира намерената команда за деинсталация от registry.
        command_text = uninstall_string.strip()
        if not command_text:
            raise RuntimeError(f"Няма команда за премахване на {display_name}.")
        normalized = command_text.lower()
        if "msiexec" in normalized and "/qn" not in normalized and "/quiet" not in normalized:
            command_text = f"{command_text} /qn /norestart"
        result = subprocess.run(
            ["cmd", "/c", command_text],
            capture_output=True,
            text=True,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        output = self._collect_command_output(result)
        self._append_command_output(output)
        if result.returncode != 0:
            raise RuntimeError(output or f"Премахването на {display_name} върна код {result.returncode}.")
        return output or f"{display_name} беше премахнат успешно."

    # Помощна функция за is winget package installed.
    def _is_winget_package_installed(self, package_id: str) -> tuple[bool, str]:
        # Проверява дали даден winget пакет вече е инсталиран.
        winget_exe = find_winget_executable()
        if not winget_exe:
            return False, ""
        result = subprocess.run(
            [winget_exe, "list", "--id", package_id, "--exact", "--accept-source-agreements"],
            capture_output=True,
            text=True,
            check=False,
            timeout=60,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        output = self._collect_command_output(result)
        normalized = output.lower()
        if "no installed package found" in normalized or "no package found matching input criteria" in normalized:
            return False, output
        if package_id.lower() in normalized:
            return True, output
        useful_lines = [
            line.strip()
            for line in output.splitlines()
            if line.strip() and "---" not in line and not line.strip().lower().startswith("name")
        ]
        return bool(useful_lines), output

    # Стартира winget uninstall command и връща резултата.
    def _run_winget_uninstall_command(self, winget_exe: str, package_id: str, label: str) -> str:
        # Премахва winget пакет преди нова инсталация.
        result = subprocess.run(
            [
                winget_exe,
                "uninstall",
                "--id",
                package_id,
                "--exact",
                "--silent",
                "--disable-interactivity",
                "--accept-package-agreements",
                "--accept-source-agreements",
            ],
            capture_output=True,
            text=True,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        output = self._collect_command_output(result)
        self._append_command_output(output)
        normalized = output.lower()
        if result.returncode != 0 and "no installed package found" not in normalized and "no package found matching input criteria" not in normalized:
            raise RuntimeError(output or f"{label} uninstall returned code {result.returncode}.")
        return output or f"{label} removal finished."

    # Помощна функция за ensure auto install vars.
    def _ensure_auto_install_vars(self, tasks: list[dict[str, str]]) -> None:
        # Подготвя променливите за тикчетата в списъка с програми.
        if not self.auto_install_vars:
            self.auto_install_vars = {task["id"]: tk.BooleanVar(value=False) for task in tasks}
        for task in tasks:
            self.auto_install_vars.setdefault(task["id"], tk.BooleanVar(value=False))
            self.auto_remove_vars.setdefault(task["id"], tk.BooleanVar(value=False))

    # Помощна функция за task supports remove.
    def _task_supports_remove(self, task: dict[str, str]) -> bool:
        # Връща дали за тази задача може първо да махнем старата версия.
        if task["type"] in {"office_offline", "office_online", "adobe"}:
            return True
        spec = self._local_task_spec(task["id"])
        if not spec:
            return False
        return spec.get("detect_mode") in {"winget", "registry"}

    # Помощна функция за task install state.
    def _task_install_state(self, task: dict[str, str]) -> tuple[bool, str]:
        # Проверява дали този софтуер вече е инсталиран.
        task_type = task["type"]
        action_id = task["id"]
        if task_type == "office_offline":
            info = self._office_install_info(action_id)
            return bool(info.installed), info.display_name or "Office пакет"
        if task_type == "office_online":
            installed, details, _ = self._office_online_install_state(action_id)
            return installed, details
            package = get_online_package(action_id)
            installed, output = self._is_winget_package_installed(package.winget_id)
            return installed, output or package.winget_id
        if task_type == "adobe":
            installed, detail, _ = self._adobe_install_state()
            return installed, detail or "Adobe Reader"
        if task_type == "local_installer":
            spec = self._local_task_spec(action_id)
            if not spec:
                return False, ""
            detect_mode = spec.get("detect_mode", "")
            detect_value = spec.get("detect_value", "")
            if detect_mode == "winget" and detect_value:
                installed, output = self._is_winget_package_installed(detect_value)
                return installed, output or detect_value
            if detect_mode == "registry" and detect_value:
                installed, detail, _ = self._find_installed_registry_app((detect_value,))
                return installed, detail
        if task_type == "resource_info":
            local_file = self._find_resource_local_file(task["resource_id"])
            return bool(local_file), str(local_file) if local_file else "Файлът още липсва"
        if task_type == "standalone_local":
            local_file = Path(task["local_path"])
            return local_file.exists(), str(local_file) if local_file.exists() else "Файлът липсва"
        return False, ""

    # Задава auto install selection според избраното действие.
    def _set_auto_install_selection(self, value: bool) -> None:
        # С едно действие маркира или изчиства всички задачи.
        for var in self.auto_install_vars.values():
            var.set(value)

    # Стартира auto install task и връща резултата.
    def _run_auto_install_task(self, task: dict[str, str]) -> str:
        # Пуска избраната задача от общия списък.
        task_type = task["type"]
        action_id = task["id"]
        remove_first = str(task.get("remove_first", "")).lower() in {"1", "true", "yes"} or bool(task.get("remove_first"))
        if task_type == "office_offline":
            return self._run_auto_office_offline(action_id, remove_first)
        if task_type == "office_online":
            return self._run_auto_office_online(action_id, remove_first)
        if task_type == "adobe":
            return self._run_auto_adobe_reader(remove_first)
        if task_type == "local_installer":
            return self._run_auto_local_installer(action_id, remove_first)
        if task_type == "resource_info":
            local_file = self._find_resource_local_file(task["resource_id"])
            if not local_file:
                raise RuntimeError("Локалният файл за този ресурс още липсва.")
            return self._run_generic_resource_task(task["label"], local_file)
        if task_type == "standalone_local":
            local_file = Path(task["local_path"])
            if not local_file.exists():
                raise RuntimeError("Локалният installer файл липсва.")
            return self._run_generic_resource_task(task["label"], local_file)
        if task_type == "language":
            return self._run_auto_language_action(action_id)
        raise ValueError(f"Непознат тип задача: {task_type}")

    # Стартира auto office offline и връща резултата.
    def _run_auto_office_offline(self, action_id: str, remove_first: bool = False) -> str:
        # Стартира offline Office инсталация и маха старата версия само ако е избрано.
        installer = get_office_offline_installer(action_id)
        office_info = detect_installed_office(action_id)
        missing_parts: list[str] = []
        if not installer.installers_root.exists():
            missing_parts.append(f"Installers папката липсва: {installer.installers_root}")
        if not installer.setup_path.exists():
            missing_parts.append(f"setup.exe липсва: {installer.setup_path}")
        if not installer.config_path.exists():
            missing_parts.append(f"Configuration файлът липсва: {installer.config_path}")
        if missing_parts:
            raise RuntimeError("\n".join(missing_parts))

        if remove_first and office_info.installed and office_info.uninstall_string:
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    35,
                    f"Премахване на стара версия за {installer.label}...",
                    office_info.display_name,
                ),
            )
            self._run_office_uninstall_command(action_id, office_info.display_name, office_info.uninstall_string)

        command = [str(installer.setup_path), "/configure", str(installer.config_path)]
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
            cwd=str(installer.setup_path.parent),
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        output = self._collect_command_output(result)
        if result.returncode != 0:
            raise RuntimeError(output or f"Office installer върна код {result.returncode}.")
        return output or f"{installer.label} завърши успешно."

    # Стартира auto office online и връща резултата.
    def _run_auto_office_online(self, action_id: str, remove_first: bool = False) -> str:
        return self._run_office_online_install_core(action_id, remove_existing=remove_first)
        # Стартира online Office инсталация и по желание маха старата версия.
        package = get_online_package(action_id)
        status = check_online_package(action_id)
        if not status.available:
            raise RuntimeError(status.message)
        winget_exe = find_winget_executable()
        if not winget_exe:
            raise RuntimeError("Winget не е открит.")

        installed_now, installed_output = self._is_winget_package_installed(package.winget_id)
        if remove_first and installed_now:
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    35,
                    f"Премахване на стара версия за {package.label}...",
                    installed_output or package.winget_id,
                ),
            )
            self._run_winget_uninstall_command(winget_exe, package.winget_id, package.label)

        command = [
            winget_exe,
            "install",
            "--id",
            package.winget_id,
            "--source",
            "winget",
            "--silent",
            "--disable-interactivity",
            "--accept-package-agreements",
            "--accept-source-agreements",
        ]
        self._append_command_output("Стартира winget online инсталация...")
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        output = self._collect_command_output(result)
        self._append_command_output(output)
        if result.returncode != 0:
            raise RuntimeError(output or f"Winget върна код {result.returncode}.")
        return output or f"{package.label} е инсталиран успешно."

    # Стартира auto adobe reader и връща резултата.
    def _run_auto_adobe_reader(self, remove_first: bool = False) -> str:
        # Стартира Adobe Reader с по-ясна проверка и по желание маха старата версия.
        winget_exe = find_winget_executable()
        if not winget_exe:
            raise RuntimeError("Winget не е открит.")

        installed_now, installed_output, uninstall_string = self._adobe_install_state()
        if remove_first and installed_now:
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    35,
                    "Премахване на стара версия на Adobe Reader...",
                    installed_output or ADOBE_READER_WINGET_ID,
                ),
            )
            winget_installed, _ = self._is_winget_package_installed(ADOBE_READER_WINGET_ID)
            if winget_installed:
                self._run_winget_uninstall_command(winget_exe, ADOBE_READER_WINGET_ID, "Adobe Reader")
            elif uninstall_string:
                self._run_uninstall_string_command("Adobe Reader", uninstall_string)

        command = [
            winget_exe,
            "install",
            "--id",
            ADOBE_READER_WINGET_ID,
            "--source",
            "winget",
            "--silent",
            "--disable-interactivity",
            "--accept-package-agreements",
            "--accept-source-agreements",
        ]
        self._append_command_output("Проверка и стартиране на Adobe Reader чрез winget...")
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        output = self._collect_command_output(result)
        self._append_command_output(output)
        if result.returncode != 0:
            raise RuntimeError(output or f"Adobe Reader инсталацията върна код {result.returncode}.")
        self.adobe_reader_status_cache = None
        return output or "Adobe Reader е инсталиран/обновен успешно."

    # Стартира auto local installer и връща резултата.
    def _run_auto_local_installer(self, action_id: str, remove_first: bool = False) -> str:
        # Стартира локален installer от Installers папката.
        spec = self._local_task_spec(action_id)
        if not spec:
            raise RuntimeError("Липсва настройка за този локален installer.")
        local_file = self._find_resource_local_file(spec["resource_id"])
        if not local_file:
            raise RuntimeError(f"Локалният installer липсва за {spec['label']}.")

        detect_mode = spec.get("detect_mode", "")
        detect_value = spec.get("detect_value", "")
        if remove_first and detect_mode == "winget" and detect_value:
            installed_now, installed_output = self._is_winget_package_installed(detect_value)
            if installed_now:
                self.root.after(
                    0,
                    lambda: self._update_activation_progress(
                        35,
                        f"Премахване на стара версия за {spec['label']}...",
                        installed_output or detect_value,
                    ),
                )
                winget_exe = find_winget_executable()
                if winget_exe:
                    self._run_winget_uninstall_command(winget_exe, detect_value, spec["label"])

        command = [str(local_file)]
        silent_args = str(spec.get("silent_args", "")).strip()
        if silent_args:
            command.extend(part for part in silent_args.split(" ") if part)
        self._append_command_output(f"Стартиране на локален installer: {local_file.name}")
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
            cwd=str(local_file.parent),
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        output = self._collect_command_output(result)
        self._append_command_output(output)
        if result.returncode != 0:
            raise RuntimeError(output or f"{spec['label']} върна код {result.returncode}.")
        return output or f"{spec['label']} стартира успешно."

    # Стартира generic resource task и връща резултата.
    def _run_generic_resource_task(self, label: str, local_file: Path) -> str:
        # Стартира общ локален .exe/.msi/.bat/.cmd ресурс от Installers папката.
        extension = local_file.suffix.lower()
        if extension not in {".exe", ".msi", ".bat", ".cmd"}:
            return f"Наличен локален ресурс: {local_file}"
        if extension == ".msi":
            command = ["msiexec", "/i", str(local_file)]
        elif extension in {".bat", ".cmd"}:
            command = ["cmd", "/c", str(local_file)]
        else:
            command = [str(local_file)]
        self._append_command_output(f"Стартиране на локален ресурс: {local_file.name}")
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
            cwd=str(local_file.parent),
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        output = self._collect_command_output(result)
        self._append_command_output(output)
        if result.returncode != 0:
            raise RuntimeError(output or f"{label} върна код {result.returncode}.")
        return output or f"{label} стартира успешно."

    # Обработва събитието on program selector mousewheel.
    def _on_program_selector_mousewheel(self, canvas: tk.Canvas, event: tk.Event) -> str:
        # Позволява скрол с мишката в прозореца за избор на програми.
        delta = getattr(event, "delta", 0)
        if getattr(event, "num", None) == 4:
            delta = 120
        elif getattr(event, "num", None) == 5:
            delta = -120
        if delta == 0:
            return "break"
        canvas.yview_scroll(-1 if delta > 0 else 1, "units")
        return "break"

    # Помощна функция за bind program selector mousewheel.
    def _bind_program_selector_mousewheel(self, widget: tk.Widget, canvas: tk.Canvas) -> None:
        # Връзва колелцето на мишката към дадения списък.
        widget.bind("<MouseWheel>", lambda event: self._on_program_selector_mousewheel(canvas, event))
        widget.bind("<Button-4>", lambda event: self._on_program_selector_mousewheel(canvas, event))
        widget.bind("<Button-5>", lambda event: self._on_program_selector_mousewheel(canvas, event))

    # Рисува program selector loading върху текущия екран.
    def _render_program_selector_loading(
        self,
        parent: tk.Widget,
        *,
        percent_var: tk.IntVar,
        status_var: tk.StringVar,
        detail_var: tk.StringVar,
    ) -> None:
        # Показва loading екран, докато се проверява наличният софтуер.
        loading_wrap = max(520, self.right_subtitle_wrap + self._scale_px(120))
        for child in parent.winfo_children():
            child.destroy()

        wrapper = tk.Frame(parent, bg="#0b1d0f")
        wrapper.pack(fill="both", expand=True, padx=18, pady=18)

        tk.Label(
            wrapper,
            text="Проверка на наличния софтуер",
            font=self._font(16, "bold", "Segoe UI Semibold"),
            bg="#0b1d0f",
            fg="#edffef",
        ).pack(anchor="center", pady=(34, 10))
        tk.Label(
            wrapper,
            textvariable=status_var,
            font=self._font(11),
            bg="#0b1d0f",
            fg="#bff3c8",
            justify="center",
            wraplength=loading_wrap,
        ).pack(anchor="center", pady=(4, 6))
        ttk.Progressbar(wrapper, maximum=100, variable=percent_var, length=520).pack(pady=(8, 10))
        tk.Label(
            wrapper,
            textvariable=detail_var,
            font=("Consolas", 10),
            bg="#0b1d0f",
            fg="#ffe08a",
            justify="center",
            wraplength=loading_wrap,
        ).pack(anchor="center", pady=(4, 0))

    # Зарежда program selector async от файл или конфигурация.
    def _load_program_selector_async(
        self,
        parent: tk.Widget,
        *,
        wraplength: int,
        start_button_text: str,
        show_close_button: bool,
        show_descriptions: bool = True,
    ) -> None:
        # Зарежда проверките във фонов режим и после рисува менюто наведнъж.
        if self.program_selector_tasks_cache and self.program_selector_status_cache and not self.program_selector_scan_running:
            self._build_program_selector_content(
                parent,
                wraplength=wraplength,
                start_button_text=start_button_text,
                show_close_button=show_close_button,
                show_descriptions=show_descriptions,
                tasks=self.program_selector_tasks_cache,
                status_map=self.program_selector_status_cache,
            )
            return

        percent_var = tk.IntVar(value=0)
        status_var = tk.StringVar(value="Подготовка на списъка с програми...")
        detail_var = tk.StringVar(value="0%")
        self._render_program_selector_loading(
            parent,
            percent_var=percent_var,
            status_var=status_var,
            detail_var=detail_var,
        )

        # Показва error в интерфейса.
        def show_error(message: str) -> None:
            # Ако проверката се счупи, показваме явна грешка вместо празно меню.
            self.program_selector_scan_running = False
            if not parent.winfo_exists():
                return
            for child in parent.winfo_children():
                child.destroy()
            wrapper = tk.Frame(parent, bg="#102515")
            wrapper.pack(fill="both", expand=True, padx=18, pady=18)
            tk.Label(
                wrapper,
                text="Автоматичният инсталатор не успя да зареди списъка",
                font=self._font(15, "bold", "Segoe UI Semibold"),
                bg="#102515",
                fg="#ffd3d3",
                justify="center",
                wraplength=wraplength,
            ).pack(anchor="center", pady=(30, 10))
            tk.Label(
                wrapper,
                text=message,
                font=self._font(10),
                bg="#102515",
                fg="#fff0f0",
                justify="left",
                wraplength=wraplength,
            ).pack(anchor="center", pady=(4, 0))

        # Помощна функция за worker.
        def worker() -> None:
            self.program_selector_scan_running = True
            try:
                tasks = self._auto_install_tasks()
                total = max(1, len(tasks))
                status_map: dict[str, tuple[bool, str]] = {}
                for index, task in enumerate(tasks, start=1):
                    status_map[task["id"]] = self._safe_task_install_state(task)
                    percent = int(index * 100 / total)
                    self.root.after(
                        0,
                        lambda percent=percent, task=task, index=index, total=total: (
                            percent_var.set(percent),
                            status_var.set(f"Проверка {index}/{total}: {task['label']}"),
                            detail_var.set(f"{percent}%"),
                        ),
                    )

                # Помощна функция за finish.
                def finish() -> None:
                    self.program_selector_scan_running = False
                    self.program_selector_tasks_cache = tasks
                    self.program_selector_status_cache = status_map
                    if not parent.winfo_exists():
                        return
                    self._build_program_selector_content(
                        parent,
                        wraplength=wraplength,
                        start_button_text=start_button_text,
                        show_close_button=show_close_button,
                        show_descriptions=show_descriptions,
                        tasks=tasks,
                        status_map=status_map,
                    )

                self.root.after(0, finish)
            except Exception as exc:
                self.root.after(0, lambda: show_error(str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    # Подготвя program selector content според избраните настройки.
    def _build_program_selector_content(
        self,
        parent: tk.Widget,
        *,
        wraplength: int,
        start_button_text: str,
        show_close_button: bool = False,
        show_descriptions: bool = True,
        tasks: list[dict[str, str]] | None = None,
        status_map: dict[str, tuple[bool, str]] | None = None,
    ) -> None:
        # Рисува общия списък с категории, тикчета и бутони за избор.
        for child in parent.winfo_children():
            child.destroy()
        tasks = tasks or self._auto_install_tasks()
        status_map = status_map or {}
        self._ensure_auto_install_vars(tasks)

        if not tasks:
            empty_frame = tk.Frame(parent, bg="#102515")
            empty_frame.pack(fill="both", expand=True, padx=18, pady=18)
            tk.Label(
                empty_frame,
                text="Няма намерени задачи за автоматичния инсталатор.",
                font=self._font(15, "bold", "Segoe UI Semibold"),
                bg="#102515",
                fg="#ffd3d3",
                justify="center",
                wraplength=wraplength,
            ).pack(anchor="center", pady=(30, 10))
            tk.Label(
                empty_frame,
                text="Списъкът е празен. Това значи, че нещо липсва в конфигурацията на задачите.",
                font=self._font(10),
                bg="#102515",
                fg="#fff0f0",
                justify="center",
                wraplength=wraplength,
            ).pack(anchor="center", pady=(4, 0))
            return

        canvas = tk.Canvas(parent, bg="#0b1d0f", highlightthickness=0, height=360)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        task_frame = tk.Frame(canvas, bg="#0b1d0f")
        task_frame.bind("<Configure>", lambda event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=task_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(16, 0), pady=8)
        scrollbar.pack(side="right", fill="y", padx=(0, 16), pady=8)
        self._bind_program_selector_mousewheel(canvas, canvas)
        self._bind_program_selector_mousewheel(task_frame, canvas)

        current_category = ""
        for task in tasks:
            installed_now, installed_text = status_map.get(task["id"], (False, ""))
            if task["category"] != current_category:
                current_category = task["category"]
                category_label = tk.Label(
                    task_frame,
                    text=current_category,
                    font=("Segoe UI Semibold", 12),
                    bg="#0b1d0f",
                    fg="#c9ffd0",
                )
                category_label.pack(anchor="w", padx=12, pady=(12, 4))
                self._bind_program_selector_mousewheel(category_label, canvas)

            row = tk.Frame(task_frame, bg="#112716", padx=10, pady=8)
            row.pack(fill="x", padx=12, pady=4)
            self._bind_program_selector_mousewheel(row, canvas)
            check_button = tk.Checkbutton(
                row,
                variable=self.auto_install_vars[task["id"]],
                bg="#112716",
                activebackground="#112716",
                selectcolor="#174327",
                fg="#edffef",
                activeforeground="#ffffff",
                text=task["label"],
                font=("Segoe UI Semibold", 10),
                anchor="w",
            )
            check_button.pack(anchor="w", fill="x")
            self._bind_program_selector_mousewheel(check_button, canvas)
            if show_descriptions:
                description_label = tk.Label(
                    row,
                    text=task["description"],
                    bg="#112716",
                    fg="#91b897",
                    font=("Segoe UI", 9),
                    wraplength=wraplength,
                    justify="left",
                )
                description_label.pack(anchor="w", padx=(24, 0), pady=(2, 0))
                self._bind_program_selector_mousewheel(description_label, canvas)

            if installed_now:
                installed_label = tk.Label(
                    row,
                    text=f"Намерено: {installed_text}",
                    bg="#112716",
                    fg="#c8f7cb",
                    font=("Segoe UI", 9),
                    justify="left",
                )
                installed_label.pack(anchor="w", padx=(24, 0), pady=(4, 0))
                self._bind_program_selector_mousewheel(installed_label, canvas)

            if self._task_supports_remove(task):
                remove_text = "Премахни старата версия и после инсталирай новата"
                remove_check = tk.Checkbutton(
                    row,
                    variable=self.auto_remove_vars[task["id"]],
                    bg="#112716",
                    activebackground="#112716",
                    selectcolor="#3e1b1b",
                    fg="#ffd9d9",
                    activeforeground="#ffffff",
                    text=remove_text,
                    font=("Segoe UI", 9),
                    anchor="w",
                )
                if not installed_now:
                    self.auto_remove_vars[task["id"]].set(False)
                    remove_check.config(state="disabled", fg="#8e7b7b")
                remove_check.pack(anchor="w", fill="x", padx=(24, 0), pady=(4, 0))
                self._bind_program_selector_mousewheel(remove_check, canvas)

        controls = tk.Frame(parent, bg="#102515")
        controls.pack(fill="x", padx=16, pady=(8, 16))
        self._bind_program_selector_mousewheel(controls, canvas)

        select_all_button = tk.Button(
            controls,
            text="Избери всичко",
            command=lambda: self._set_auto_install_selection(True),
            font=("Segoe UI Semibold", 10),
            bg="#1f6fb2",
            fg="#f4fbff",
            activebackground="#2b8ddd",
            activeforeground="#ffffff",
            bd=0,
            padx=16,
            pady=9,
            cursor="hand2",
        )
        select_all_button.pack(side="left", padx=(0, 8))
        self._bind_program_selector_mousewheel(select_all_button, canvas)
        clear_button = tk.Button(
            controls,
            text="Изчисти",
            command=lambda: self._set_auto_install_selection(False),
            font=("Segoe UI Semibold", 10),
            bg="#36403a",
            fg="#d8e2db",
            activebackground="#465049",
            activeforeground="#ffffff",
            bd=0,
            padx=16,
            pady=9,
            cursor="hand2",
        )
        clear_button.pack(side="left")
        self._bind_program_selector_mousewheel(clear_button, canvas)

        if show_close_button:
            close_button = tk.Button(
                controls,
                text="Затвори",
                command=self._close_program_selector_window,
                font=("Segoe UI Semibold", 10),
                bg="#5a2424",
                fg="#fff0f0",
                activebackground="#7b3030",
                activeforeground="#ffffff",
                bd=0,
                padx=16,
                pady=9,
                cursor="hand2",
            )
            close_button.pack(side="right", padx=(8, 0))
            self._bind_program_selector_mousewheel(close_button, canvas)

        start_button = tk.Button(
            controls,
            text=start_button_text,
            command=self._start_auto_installer,
            font=("Segoe UI Semibold", 11),
            bg="#1f8f43",
            fg="#f5fff7",
            activebackground="#28b155",
            activeforeground="#ffffff",
            bd=0,
            padx=20,
            pady=10,
            cursor="hand2",
        )
        start_button.pack(side="right")
        self._bind_program_selector_mousewheel(start_button, canvas)

    # Помощна функция за close program selector window.
    def _close_program_selector_window(self) -> None:
        # Затваря отделния прозорец за избор на програми.
        if self.program_selector_window and self.program_selector_window.winfo_exists():
            self.program_selector_window.destroy()
        self.program_selector_window = None

    # Помощна функция за safe task install state.
    def _safe_task_install_state(self, task: dict[str, str]) -> tuple[bool, str]:
        # Ako nqkoq proverka grymne na nov komputar, spisykyt pak ostava viden.
        cached = self.program_selector_status_cache.get(task["id"])
        if cached is not None:
            return cached
        try:
            return self._task_install_state(task)
        except Exception as exc:
            return False, f"Статусът не може да се прочете: {exc}"

    # Отваря program selector window или съответния прозорец.
    def _open_program_selector_window(self) -> None:
        # Отваря нов прозорец с пълния избор от програми за инсталиране.
        if self.program_selector_window and self.program_selector_window.winfo_exists():
            self.program_selector_window.focus_force()
            return

        window = tk.Toplevel(self.root)
        self.program_selector_window = window
        window.title("Избор на програми")
        window.configure(bg="#102515")
        window.minsize(920, 700)
        window.transient(self.root)
        apply_app_icon(window)
        self._center_window(window, 980, 720)
        window.protocol("WM_DELETE_WINDOW", self._close_program_selector_window)

        outer = tk.Frame(
            window,
            bg="#102515",
            bd=0,
            highlightthickness=1,
            highlightbackground="#2d7f4a",
        )
        outer.pack(fill="both", expand=True, padx=12, pady=12)

        header = tk.Frame(outer, bg="#102515")
        header.pack(fill="x", padx=16, pady=(14, 8))
        tk.Label(
            header,
            text="Избор на програми",
            font=("Segoe UI Semibold", 16),
            bg="#102515",
            fg="#edffef",
        ).pack(anchor="w")
        tk.Label(
            header,
            text="Тук можеш с тикче да отбележиш всичко, което искаш да се инсталира. След това задачите ще се изпълнят една след друга.",
            font=self._font(10),
            bg="#102515",
            fg="#9bc39e",
            wraplength=880,
            justify="left",
        ).pack(anchor="w", pady=(4, 0))

        content_holder = tk.Frame(outer, bg="#102515")
        content_holder.pack(fill="both", expand=True)
        self._load_program_selector_async(
            content_holder,
            wraplength=860,
            start_button_text="Инсталирай избраните",
            show_close_button=True,
        )

    # Рисува auto installer върху текущия екран.
    def _render_auto_installer(self) -> None:
        # Това е обновената версия на страницата за автоматичен инсталатор.
        for index in range(12):
            self.cards_frame.rowconfigure(index, weight=0, minsize=0, uniform="")
            self.cards_frame.columnconfigure(index, weight=0, minsize=0, uniform="")
        self.cards_frame.columnconfigure(0, weight=1)
        self.cards_frame.rowconfigure(0, weight=1)
        self.cards_frame.update_idletasks()
        frame_width = self.cards_frame.winfo_width()
        frame_height = self.cards_frame.winfo_height()
        available_width = max(720, frame_width - self._scale_px(48)) if frame_width > 1 else max(720, self.root.winfo_width() - self.sidebar_width - self._scale_px(130))
        available_height = max(520, frame_height - self._scale_px(48)) if frame_height > 1 else max(520, self.root.winfo_height() - self.header_height_px - self._scale_px(170))
        panel_width = min(self._scale_px(920), available_width)
        panel_height = min(self._scale_px(650), available_height)
        selector_wrap = max(520, panel_width - self._scale_px(90))

        outer = tk.Frame(
            self.cards_frame,
            bg="#102515",
            bd=0,
            highlightthickness=1,
            highlightbackground="#2d7f4a",
            width=panel_width,
            height=panel_height,
        )
        outer.place(relx=0.5, rely=0.5, anchor="center", width=panel_width, height=panel_height)
        outer.pack_propagate(False)

        header = tk.Frame(outer, bg="#102515")
        header.pack(fill="x", padx=16, pady=(14, 4))
        tk.Label(
            header,
            text="Автоматичен инсталатор",
            font=("Segoe UI Semibold", 16),
            bg="#102515",
            fg="#edffef",
        ).pack(anchor="center")

        content_holder = tk.Frame(outer, bg="#102515")
        content_holder.pack(fill="both", expand=True)
        self._load_program_selector_async(
            content_holder,
            wraplength=selector_wrap,
            start_button_text="Инсталирай избраните",
            show_close_button=False,
            show_descriptions=False,
        )

        self.page_label.config(text="Автоматичен режим")
        self.prev_button.config(state="disabled")
        self.next_button.config(state="disabled")
        self.back_button.config(state="normal" if self.history else "disabled")

    # Помощна функция за center window.
    def _center_window(self, window: tk.Toplevel, width: int, height: int) -> None:
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        safe_width = min(width, max(420, screen_width - 40))
        safe_height = min(height, max(260, screen_height - 60))
        center_geometry(window, safe_width, safe_height)

    # Помощна функция за scale px.
    def _scale_px(self, value: int | float) -> int:
        # Preobrazuva pixel stoinosti според tekushtiq UI scale.
        return max(1, int(round(float(value) * self.ui_scale)))

    # Помощна функция за font.
    def _font(self, size: int, weight: str = "", family: str = "Segoe UI") -> tuple[str, int] | tuple[str, int, str]:
        # Vrushta ednakvo skaliiран font za celiq glaven ekran.
        scaled_size = max(8, int(round(size * self.ui_scale)))
        if weight:
            return (family, scaled_size, weight)
        return (family, scaled_size)

    # Обновява layout metrics след промяна в състоянието.
    def _update_layout_metrics(self) -> None:
        # Smята osnovnite layout meri според DPI, rezolyuciqta i tekushtata shirina na prozoreca.
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        current_width = self.root.winfo_width() or screen_width
        current_height = self.root.winfo_height() or screen_height
        try:
            dpi = float(self.root.winfo_fpixels("1i"))
        except Exception:
            dpi = BASE_DPI

        dpi_scale = dpi / BASE_DPI
        resolution_scale = min(screen_width / 1600.0, screen_height / 900.0)
        width_scale = min(current_width / 1280.0, current_height / 840.0)
        self.ui_scale = clamp(min(dpi_scale, 1.18) * clamp(min(resolution_scale, width_scale), 0.88, 1.16), 0.88, 1.18)

        if current_width < 1250:
            self.ui_scale = min(self.ui_scale, 0.96)
        if current_width < 1100:
            self.ui_scale = min(self.ui_scale, 0.90)

        self.sidebar_width = max(250, min(360, self._scale_px(320)))
        self.header_height_px = max(78, self._scale_px(90))
        self.header_title_size = max(18, self._scale_px(22))
        self.header_subtitle_size = max(9, self._scale_px(10))
        self.body_text_size = max(8, self._scale_px(9))
        self.button_text_size = max(9, self._scale_px(10))
        self.section_title_size = max(12, self._scale_px(15))
        self.card_title_size = max(10, self._scale_px(12))
        self.card_desc_size = max(8, self._scale_px(9))
        self.language_panel_width = max(240, self._scale_px(280))
        self.language_status_wrap = max(200, self._scale_px(230))
        self.system_info_wrap = max(220, self._scale_px(280))
        self.resource_wrap = max(360, self._scale_px(520))
        self.right_subtitle_wrap = max(420, self._scale_px(630))
        self.content_pad_x = max(12, self._scale_px(20))
        self.content_pad_y = max(12, self._scale_px(18))
        self.nav_button_char_width = 10 if current_width < 1200 else 11
        self.card_button_width_px = max(210, self._scale_px(CARD_BUTTON_PIXEL_WIDTH))
        self.card_button_height_px = max(42, self._scale_px(CARD_BUTTON_PIXEL_HEIGHT))
        self.card_action_gap_px = max(6, self._scale_px(8))
        self.card_title_wrap = max(240, min(420, current_width // 3 - 70))
        self.card_desc_wrap = max(240, min(420, current_width // 3 - 70))
        self.compact_card_title_wrap = max(260, min(440, current_width // 3 - 50))
        self.compact_card_desc_wrap = max(260, min(440, current_width // 3 - 50))
        self.scaled_card_min_height = max(170, self._scale_px(CARD_MIN_HEIGHT))
        self.scaled_menu_card_min_height = {
            key: max(self.scaled_card_min_height, self._scale_px(value))
            for key, value in MENU_CARD_MIN_HEIGHT.items()
        }

    # Помощна функция за apply responsive theme.
    def _apply_responsive_theme(self) -> None:
        # Obnovqva osnovnite widget-Рё sled premervane, za da stoqt dobre na razlichni monitori.
        self.header.configure(height=self.header_height_px)
        self.title_label.configure(font=self._font(22, "bold", "Segoe UI Semibold"))
        self.subtitle_label.configure(font=self._font(10), wraplength=max(420, self._scale_px(720)))
        self.header_device_chip.configure(font=self._font(9, "bold", "Segoe UI Semibold"))
        self.version_chip.configure(font=self._font(9, "bold", "Segoe UI Semibold"))
        self.header_admin_chip.configure(font=self._font(9, "bold", "Segoe UI Semibold"))
        self.header_exit_button.configure(font=self._font(10, "bold", "Segoe UI Semibold"), width=max(8, int(10 * self.ui_scale)))
        self.header_dashboard_button.configure(font=self._font(9, "bold", "Segoe UI Semibold"), width=max(20, int(22 * self.ui_scale)))
        self.header_exit_button.place_configure(x=-24, y=self._scale_px(22))
        version_x = max(420, self._scale_px(520))
        admin_x = max(500, self._scale_px(608))
        self.version_chip.place_configure(x=version_x, y=self._scale_px(22))
        self.header_admin_chip.place_configure(x=admin_x, y=self._scale_px(22))
        self.header.update_idletasks()
        self.header_device_chip.place_configure(x=26, y=self._scale_px(56))
        self.left_panel.configure(width=self.sidebar_width)
        self.menu_title.configure(font=self._font(15, "bold", "Segoe UI Semibold"))
        self.menu_path.configure(font=self._font(10), wraplength=self.system_info_wrap)
        self.sidebar_section_label.configure(font=self._font(8, "bold", "Segoe UI Semibold"))
        self.sidebar_toggle_label.configure(font=self._font(15, "normal", "Segoe UI Symbol"))
        self.sidebar_clock_card.winfo_children()[0].configure(font=self._font(16, "bold", "Segoe UI Semibold"))
        self.sidebar_clock_card.winfo_children()[1].configure(font=self._font(9))
        for parts in self.sidebar_nav_buttons.values():
            parts["title"].configure(font=self._font(9, "bold", "Segoe UI Semibold"))
            parts["subtitle"].configure(font=self._font(7), wraplength=max(150, self.sidebar_width - 110))
            parts["arrow"].configure(font=self._font(12, "bold", "Segoe UI Semibold"))
        self.system_info.configure(font=("Consolas", max(8, self.body_text_size)), wraplength=self.system_info_wrap)
        self.hint_label.configure(font=self._font(9), wraplength=self.system_info_wrap)
        self.health_title.configure(font=self._font(14, "bold", "Segoe UI Semibold"))
        if self.health_loading_label.winfo_exists():
            self.health_loading_label.configure(font=self._font(10), wraplength=max(220, self.sidebar_width - 60))
        self.status_bar.configure(font=self._font(10))
        self.card_title.configure(font=self._font(19, "bold", "Segoe UI Semibold"))
        self.card_subtitle.configure(font=self._font(10), wraplength=self.right_subtitle_wrap)
        self.resource_title.configure(font=self._font(11, "bold", "Segoe UI Semibold"))
        self.resource_status_label.configure(font=self._font(9), wraplength=self.resource_wrap)
        self.resource_download_button.configure(font=self._font(9, "bold", "Segoe UI Semibold"))
        if hasattr(self, "software_summary_download_button"):
            self.software_summary_download_button.configure(font=self._font(9, "bold", "Segoe UI Semibold"))
        self.resource_details_button.configure(font=self._font(9, "bold", "Segoe UI Semibold"))
        self.update_icon_label.configure(font=self._font(16, "bold", "Segoe UI Semibold"))
        self.update_message_label.configure(font=self._font(10), wraplength=self.resource_wrap)
        self.update_action_button.configure(font=self._font(9, "bold", "Segoe UI Semibold"))
        self.update_history_button.configure(font=self._font(9, "bold", "Segoe UI Semibold"))
        self.page_label.configure(font=self._font(10))
        self.language_status_panel.configure(width=self.language_panel_width)
        self.language_status_title.configure(font=self._font(13, "bold", "Segoe UI Semibold"))
        self.language_status_label.configure(font=self._font(9), wraplength=self.language_status_wrap)
        for child in self.overview_frame.winfo_children():
            for nested in child.winfo_children():
                for widget in nested.winfo_children():
                    if isinstance(widget, tk.Label):
                        current_font = str(widget.cget("font"))
                        if "Semibold" in current_font or "bold" in current_font.lower():
                            widget.configure(font=self._font(13, "bold", "Segoe UI Semibold"))
                        else:
                            widget.configure(font=self._font(9))
        self.content.pack_configure(padx=self.content_pad_x, pady=self.content_pad_y)

    # Обработва събитието on root resize.
    def _on_root_resize(self, event: tk.Event[tk.Misc]) -> None:
        # Pri smqna na razmera preizchislyava kolonite samo sled kratko izchakvane, za da ne primigva.
        if event.widget is not self.root:
            return
        self._update_layout_metrics()
        self._apply_responsive_theme()
        width_bucket = max(1, event.width // 120)
        height_bucket = max(1, event.height // 90)
        new_bucket = (width_bucket, height_bucket)
        if new_bucket == self.last_layout_bucket:
            return
        self.last_layout_bucket = new_bucket
        if self.resize_render_job:
            self.root.after_cancel(self.resize_render_job)
        self.resize_render_job = self.root.after(120, self._rerender_after_resize)

    # Помощна функция за rerender after resize.
    def _rerender_after_resize(self) -> None:
        # Osvejava kartite sled resize, za da se vidi po-dobre tekstut na razlichni monitori.
        self.resize_render_job = None
        self._render_cards()

    # Помощна функция за choose office version.
    def _choose_office_version(self, title: str) -> str | None:
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.configure(bg="#0b1d0f")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        self._center_window(dialog, 430, 300)

        selected_version = tk.StringVar(value="")

        wrapper = tk.Frame(dialog, bg="#0b1d0f", padx=20, pady=18)
        wrapper.pack(fill="both", expand=True)

        tk.Label(
            wrapper,
            text="Choose Office version",
            font=("Segoe UI Semibold", 16),
            fg="#d9ffe0",
            bg="#0b1d0f",
        ).pack(anchor="w", pady=(0, 6))

        tk.Label(
            wrapper,
            text="The key will be saved and used only for the selected version.",
            font=("Segoe UI", 10),
            fg="#9dc7a4",
            bg="#0b1d0f",
            wraplength=380,
            justify="left",
        ).pack(anchor="w", pady=(0, 14))

        # Помощна функция за choose.
        def choose(action_id: str) -> None:
            selected_version.set(action_id)
            dialog.destroy()

        for action_id in OFFICE_ACTION_IDS:
            tk.Button(
                wrapper,
                text=get_office_version_label(action_id),
                command=lambda current=action_id: choose(current),
                font=("Segoe UI Semibold", 10),
                bg="#174327",
                fg="#eefef1",
                activebackground="#236039",
                activeforeground="#ffffff",
                bd=0,
                padx=16,
                pady=10,
                width=28,
                cursor="hand2",
            ).pack(fill="x", pady=5)

        tk.Button(
            wrapper,
            text="Cancel",
            command=dialog.destroy,
            font=("Segoe UI Semibold", 10),
            bg="#4c1c1c",
            fg="#fff4f4",
            activebackground="#7a1f1f",
            activeforeground="#ffffff",
            bd=0,
            padx=16,
            pady=10,
            width=28,
            cursor="hand2",
        ).pack(fill="x", pady=(12, 0))

        self.root.wait_window(dialog)
        return selected_version.get() or None

    # Записва windows product key за следващо използване.
    def _save_windows_product_key(self, version_label: str, store_key: str) -> None:
        # Записва ключ за избраната версия на Windows.
        existing_key = self.secure_store.get(store_key, "")
        prompt = f"Enter the {version_label} product key to save for your admin workflow:"
        product_key = ask_product_key(self.root, f"Save {version_label} Key", prompt, initialvalue=existing_key)
        if product_key is None:
            self.status_var.set(f"Saving {version_label} product key was canceled.")
            return

        normalized_key = normalize_product_key_input(product_key)
        if not normalized_key:
            messagebox.showwarning("Missing Key", "Please enter a product key before saving.", parent=self.root)
            self.status_var.set(f"No {version_label} product key was saved.")
            return

        self.secure_store[store_key] = normalized_key
        try:
            save_secure_store(self.secure_store)
        except OSError as exc:
            messagebox.showerror(
                "Save Failed",
                f"The {version_label} product key could not be saved.\n\n{exc}",
                parent=self.root,
            )
            self.status_var.set(f"Saving the {version_label} product key failed.")
            return
        self.status_var.set(f"{version_label} product key saved successfully.")
        messagebox.showinfo(
            "Key Saved",
            f"The {version_label} product key has been saved.\n\nSecure file:\n{SECURE_STORE_FILE}",
            parent=self.root,
        )

    # Записва windows10 key за следващо използване.
    def _save_windows10_key(self) -> None:
        self._save_windows_product_key("Windows 10", "windows10_product_key")

    # Записва windows11 key за следващо използване.
    def _save_windows11_key(self) -> None:
        self._save_windows_product_key("Windows 11", "windows11_product_key")

    # Показва windows product key в интерфейса.
    def _show_windows_product_key(self, version_label: str, store_key: str) -> None:
        # Показва записания ключ за избраната версия на Windows.
        saved_key = self.secure_store.get(store_key, "").strip()
        if not saved_key:
            messagebox.showinfo("No Saved Key", f"There is no saved {version_label} product key yet.", parent=self.root)
            self.status_var.set(f"No {version_label} product key is currently stored.")
            return

        messagebox.showinfo(f"Saved {version_label} Key", saved_key, parent=self.root)
        self.status_var.set(f"Displayed the saved {version_label} product key.")

    # Показва windows10 key в интерфейса.
    def _show_windows10_key(self) -> None:
        self._show_windows_product_key("Windows 10", "windows10_product_key")

    # Показва windows11 key в интерфейса.
    def _show_windows11_key(self) -> None:
        self._show_windows_product_key("Windows 11", "windows11_product_key")

    # Помощна функция за clear windows product key.
    def _clear_windows_product_key(self, version_label: str, store_key: str) -> None:
        # Изтрива записания ключ за избраната версия на Windows.
        if store_key not in self.secure_store:
            self.status_var.set(f"There is no saved {version_label} product key to remove.")
            return

        confirmed = messagebox.askyesno(
            "Clear Saved Key",
            f"Do you want to remove the saved {version_label} product key?",
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set(f"Saved {version_label} product key was kept.")
            return

        self.secure_store.pop(store_key, None)
        try:
            save_secure_store(self.secure_store)
        except OSError as exc:
            messagebox.showerror(
                "Remove Failed",
                f"The saved {version_label} product key could not be removed.\n\n{exc}",
                parent=self.root,
            )
            self.status_var.set(f"Removing the saved {version_label} product key failed.")
            return
        self.status_var.set(f"Saved {version_label} product key removed.")

    # Помощна функция за clear windows10 key.
    def _clear_windows10_key(self) -> None:
        self._clear_windows_product_key("Windows 10", "windows10_product_key")

    # Помощна функция за clear windows11 key.
    def _clear_windows11_key(self) -> None:
        self._clear_windows_product_key("Windows 11", "windows11_product_key")

    # Записва office key за следващо използване.
    def _save_office_key(self) -> None:
        selected_action = self._choose_office_version("Save Office Key")
        if not selected_action:
            self.status_var.set("Saving Office product key was canceled.")
            return

        version_label = get_office_version_label(selected_action)
        store_key = f"{selected_action}_product_key"
        existing_key = self.secure_store.get(store_key, "")
        product_key = ask_product_key(
            self.root,
            "Save Office Key",
            f"Enter the product key for {version_label}:",
            initialvalue=existing_key,
        )
        if product_key is None:
            self.status_var.set(f"Saving {version_label} product key was canceled.")
            return

        normalized_key = normalize_product_key_input(product_key)
        if not normalized_key:
            messagebox.showwarning("Missing Key", f"Please enter a product key for {version_label} before saving.", parent=self.root)
            self.status_var.set(f"No {version_label} product key was saved.")
            return

        self.secure_store[store_key] = normalized_key
        try:
            save_secure_store(self.secure_store)
        except OSError as exc:
            messagebox.showerror(
                "Save Failed",
                f"The {version_label} product key could not be saved.\n\n{exc}",
                parent=self.root,
            )
            self.status_var.set(f"Saving the {version_label} product key failed.")
            return

        self.status_var.set(f"{version_label} product key saved successfully.")
        messagebox.showinfo(
            "Key Saved",
            f"The product key for {version_label} has been saved.\n\nSecure file:\n{SECURE_STORE_FILE}",
            parent=self.root,
        )

    # Показва office key в интерфейса.
    def _show_office_key(self) -> None:
        selected_action = self._choose_office_version("Show Office Key")
        if not selected_action:
            self.status_var.set("Showing Office product key was canceled.")
            return

        version_label = get_office_version_label(selected_action)
        saved_key = self.secure_store.get(f"{selected_action}_product_key", "").strip()
        if not saved_key:
            messagebox.showinfo("No Saved Key", f"There is no saved product key for {version_label} yet.", parent=self.root)
            self.status_var.set(f"No {version_label} product key is currently stored.")
            return

        messagebox.showinfo(f"Saved {version_label} Key", saved_key, parent=self.root)
        self.status_var.set(f"Displayed the saved {version_label} product key.")

    # Помощна функция за clear office key.
    def _clear_office_key(self) -> None:
        selected_action = self._choose_office_version("Clear Office Key")
        if not selected_action:
            self.status_var.set("Removing Office product key was canceled.")
            return

        version_label = get_office_version_label(selected_action)
        store_key = f"{selected_action}_product_key"
        if store_key not in self.secure_store:
            self.status_var.set(f"There is no saved {version_label} product key to remove.")
            return

        confirmed = messagebox.askyesno(
            "Clear Saved Key",
            f"Do you want to remove the saved product key for {version_label}?",
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set(f"Saved {version_label} product key was kept.")
            return

        self.secure_store.pop(store_key, None)
        try:
            save_secure_store(self.secure_store)
        except OSError as exc:
            messagebox.showerror(
                "Remove Failed",
                f"The saved {version_label} product key could not be removed.\n\n{exc}",
                parent=self.root,
            )
            self.status_var.set(f"Removing the saved {version_label} product key failed.")
            return

        self.status_var.set(f"Saved {version_label} product key removed.")

    # Помощна функция за activate office version.
    def _activate_office_version(self, action_id: str) -> None:
        version_label = get_office_version_label(action_id)
        saved_key = self.secure_store.get(f"{action_id}_product_key", "").strip()
        if not saved_key:
            messagebox.showwarning(
                "Missing Key",
                f"Save a product key for {version_label} first, then run Office activation.",
                parent=self.root,
            )
            self.status_var.set(f"{version_label} activation could not start because no key is saved for that version.")
            return

        confirmed = messagebox.askyesno(
            f"Activate {version_label}",
            f"Run {version_label} activation now using the saved Office product key?",
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set(f"{version_label} activation was canceled.")
            return

        self.status_var.set(f"Activating {version_label}...")
        self._open_activation_window(
            title=f"{version_label} Activation Progress",
            heading=f"{version_label} Activation",
            intro=f"The application is running the saved {version_label} activation workflow.",
        )
        threading.Thread(
            target=self._run_office_activation,
            args=(version_label, saved_key),
            daemon=True,
        ).start()

    # Помощна функция за activate windows version.
    def _activate_windows_version(self, version_label: str, store_key: str) -> None:
        # Стартира активация за избраната версия на Windows с вече записания ключ.
        saved_key = self.secure_store.get(store_key, "").strip()
        if not saved_key:
            messagebox.showwarning(
                "Missing Key",
                f"Save a {version_label} product key first, then run activation.",
                parent=self.root,
            )
            self.status_var.set(f"{version_label} activation could not start because no key is saved.")
            return

        confirmed = messagebox.askyesno(
            f"Activate {version_label}",
            f"Run {version_label} activation now using the saved product key?",
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set(f"{version_label} activation was canceled.")
            return

        self.status_var.set(f"Activating {version_label}...")
        self._open_activation_window(
            title=f"{version_label} Activation Progress",
            heading=f"{version_label} Activation",
            intro="The application is running the saved activation workflow.",
        )
        threading.Thread(target=self._run_windows_activation, args=(version_label, saved_key), daemon=True).start()

    # Помощна функция за activate windows10.
    def _activate_windows10(self) -> None:
        self._activate_windows_version("Windows 10", "windows10_product_key")

    # Помощна функция за activate windows11.
    def _activate_windows11(self) -> None:
        self._activate_windows_version("Windows 11", "windows11_product_key")

    # Намира onedrive executable.
    def _find_onedrive_executable(self) -> str | None:
        candidates = [
            Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "OneDrive" / "OneDrive.exe",
            Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Microsoft OneDrive" / "OneDrive.exe",
            Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Microsoft OneDrive" / "OneDrive.exe",
        ]
        for candidate in candidates:
            if candidate.exists():
                return str(candidate)

        path_candidate = shutil.which("OneDrive.exe")
        return path_candidate

    # Помощна функция за reset onedrive.
    def _reset_onedrive(self, action_id: str) -> None:
        onedrive_exe = self._find_onedrive_executable()
        if not onedrive_exe:
            messagebox.showerror(
                "OneDrive Not Found",
                "OneDrive.exe was not found on this computer.",
                parent=self.root,
            )
            self.status_var.set("OneDrive reset could not start because OneDrive.exe was not found.")
            return

        method_label = {
            "reset_onedrive_1": "Method 1",
            "reset_onedrive_2": "Method 2",
            "reset_onedrive_3": "Method 3",
        }[action_id]

        if action_id == "reset_onedrive_3":
            confirmed = messagebox.askyesno(
                "Confirm OneDrive Reset",
                "Method 3 will remove the local OneDrive app data folder and then start OneDrive again.\n\nContinue?",
                parent=self.root,
            )
            if not confirmed:
                self.status_var.set("OneDrive reset was canceled.")
                return

        self.status_var.set(f"Resetting OneDrive with {method_label}...")
        self._open_activation_window(
            title=f"Reset OneDrive - {method_label}",
            heading=f"Reset OneDrive {method_label}",
            intro="Изпълнява се избраният метод за нулиране на OneDrive.",
        )
        threading.Thread(
            target=self._run_onedrive_reset,
            args=(action_id, onedrive_exe),
            daemon=True,
        ).start()

    # Стартира onedrive reset и връща резултата.
    def _run_onedrive_reset(self, action_id: str, onedrive_exe: str) -> None:
        local_onedrive_dir = str(Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "OneDrive")
        method_steps: dict[str, list[tuple[int, str, str]]] = {
            "reset_onedrive_1": [
                (35, "Подготвяне на стандартен reset...", f'& "{onedrive_exe}" /reset'),
                (100, "Метод 1 приключи.", "Стандартният reset на OneDrive беше изпълнен."),
            ],
            "reset_onedrive_2": [
                (35, "Спиране на процеса OneDrive...", 'Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue'),
                (80, "Стартиране на OneDrive отново...", f'Start-Process -FilePath "{onedrive_exe}"'),
                (100, "Метод 2 приключи.", "OneDrive беше спрян и стартиран отново."),
            ],
            "reset_onedrive_3": [
                (25, "Спиране на процеса OneDrive...", 'Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue'),
                (60, "Изтриване на локалната папка на OneDrive...", f'Remove-Item -Recurse -Force "{local_onedrive_dir}" -ErrorAction SilentlyContinue'),
                (90, "Стартиране на OneDrive наново...", f'Start-Process -FilePath "{onedrive_exe}"'),
                (100, "Метод 3 приключи.", "Локалните OneDrive данни бяха изчистени и клиентът беше стартиран наново."),
            ],
        }

        steps = method_steps[action_id]
        output_lines: list[str] = []
        try:
            for index, (progress_value, status_text, command) in enumerate(steps, start=1):
                self.root.after(
                    0,
                    lambda value=progress_value, step=status_text, cmd=command, step_index=index: self._update_activation_progress(
                        value,
                        step,
                        f"Стъпка {step_index}: {cmd}" if value < 100 else cmd,
                    ),
                )
                if progress_value == 100:
                    continue

                result = subprocess.run(
                    ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command],
                    capture_output=True,
                    text=True,
                    check=False,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                if result.stdout.strip():
                    output_lines.append(result.stdout.strip())
                    self.root.after(0, lambda text=result.stdout.strip(): self._append_activation_log(text))
                if result.stderr.strip():
                    output_lines.append(result.stderr.strip())
                    self.root.after(0, lambda text=result.stderr.strip(): self._append_activation_log(text))
                if result.returncode != 0:
                    raise RuntimeError("\n\n".join(output_lines) or "OneDrive reset command failed.")
        except Exception as exc:
            self.root.after(0, lambda: self._show_activation_result(False, str(exc), "OneDrive"))
            return

        final_message = output_lines[-1] if output_lines else "Операцията за Reset OneDrive завърши успешно."
        self.root.after(0, lambda: self._show_activation_result(True, final_message, "OneDrive"))

    # Обработва събитието handle language action.
    def _handle_language_action(self, action_id: str, label: str) -> None:
        if action_id == "language_refresh":
            self._refresh_language_status(show_dialog=True)
            return

        try:
            language_status = self._language_status()
            action_title, script = build_language_action(action_id, language_status)
        except Exception as exc:
            messagebox.showerror("Language Action Failed", str(exc), parent=self.root)
            self.status_var.set("Language action could not start.")
            return

        confirmed = messagebox.askyesno(
            "Language Manager",
            f"Run action now?\n\n{action_title}",
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set(f"{label} was canceled.")
            return

        self.status_var.set(f"Running {action_title}...")
        self._open_activation_window(
            title="Language Manager",
            heading=action_title,
            intro="The application is applying the selected Windows language or keyboard layout action.",
        )
        threading.Thread(
            target=self._run_language_action,
            args=(action_title, script),
            daemon=True,
        ).start()

    # Помощна функция за refresh language status.
    def _refresh_language_status(self, show_dialog: bool = False) -> None:
        self.status_var.set("Refreshing language status...")
        self._reset_language_status_cache()
        try:
            status = self._language_status()
        except Exception as exc:
            messagebox.showerror("Language Status Failed", str(exc), parent=self.root)
            self.status_var.set("Language status refresh failed.")
            if self.current_menu == "language":
                self._render_cards()
            return

        self.status_var.set("Language status refreshed.")
        self._apply_language_status_summary(
            self._build_language_status_summary(status),
            "#9aff9f" if status.has_language_pack or status.has_bulgarian else "#ffb0a8",
        )
        if self.current_menu == "language":
            self._render_cards()
        if show_dialog:
            summary = (
                f"Bulgarian added: {'Yes' if status.has_bulgarian else 'No'}\n"
                f"Language pack: {'Yes' if status.has_language_pack else 'No'}\n"
                f"BDS: {'Yes' if status.has_bds else 'No'}\n"
                f"Phonetic: {'Yes' if status.has_phonetic else 'No'}\n"
                f"Traditional: {'Yes' if status.has_traditional else 'No'}"
            )
            messagebox.showinfo("Language Status", summary, parent=self.root)

    # Стартира language action и връща резултата.
    def _run_language_action(self, action_title: str, script: str) -> None:
        output_lines: list[str] = []
        try:
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    20,
                    "Checking current language configuration...",
                    action_title,
                ),
            )
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    55,
                    "Applying Windows language command...",
                    f"PowerShell action: {action_title}",
                ),
            )
            result = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
                capture_output=True,
                text=True,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if result.stdout.strip():
                output_lines.append(result.stdout.strip())
                self.root.after(0, lambda text=result.stdout.strip(): self._append_activation_log(text))
            if result.stderr.strip():
                output_lines.append(result.stderr.strip())
                self.root.after(0, lambda text=result.stderr.strip(): self._append_activation_log(text))
            if result.returncode != 0:
                raise RuntimeError("\n\n".join(output_lines) or f"{action_title} returned code {result.returncode}.")

            self._reset_language_status_cache()
            refreshed_status = self._language_status()
            summary = (
                f"Bulgarian added: {'Yes' if refreshed_status.has_bulgarian else 'No'}\n"
                f"Language pack: {'Yes' if refreshed_status.has_language_pack else 'No'}\n"
                f"BDS: {'Yes' if refreshed_status.has_bds else 'No'}\n"
                f"Phonetic: {'Yes' if refreshed_status.has_phonetic else 'No'}\n"
                f"Traditional: {'Yes' if refreshed_status.has_traditional else 'No'}"
            )
            final_message = "\n\n".join(output_lines + [summary]) if output_lines else summary
            self.root.after(0, lambda: self._finish_language_action(action_title, True, final_message))
        except Exception as exc:
            self.root.after(0, lambda: self._finish_language_action(action_title, False, str(exc)))

    # Помощна функция за finish language action.
    def _finish_language_action(self, subject: str, success: bool, details: str) -> None:
        self._reset_language_status_cache()
        self._show_activation_result(success, details, subject)
        self._load_language_status_async()
        if self.current_menu == "language":
            self._render_cards()

    # Обработва събитието handle driver backup action.
    def _handle_driver_backup_action(self, action_id: str) -> None:
        if action_id == "driver_backup_clean":
            self._start_driver_backup(mode="clean", base_dir=desktop_path(), subject="Driver Backup (Clean)", zip_mode="keep")
            return
        if action_id == "driver_backup_full":
            self._start_driver_backup(mode="full", base_dir=desktop_path(), subject="Driver Backup (Full)", zip_mode="keep")
            return
        if action_id == "driver_recovery_usb":
            self._create_driver_recovery_usb()
            return
        if action_id == "driver_pc_report":
            self._generate_driver_pc_report()
            return
        if action_id == "driver_backup_advanced":
            self._run_driver_backup_advanced()
            return
        if action_id == "driver_restore_last":
            self._restore_drivers_from_last_backup()
            return

    # Помощна функция за start driver backup.
    def _start_driver_backup(self, mode: str, base_dir: Path, subject: str, zip_mode: str) -> None:
        confirmed = messagebox.askyesno(
            "Driver Backup",
            f"Start {subject} now?\n\nDestination base: {base_dir}",
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set(f"{subject} was canceled.")
            return

        self.status_var.set(f"Starting {subject}...")
        self._open_activation_window(
            title=subject,
            heading=subject,
            intro="The application is exporting drivers, creating logs and preparing restore information.",
        )
        threading.Thread(
            target=self._run_driver_backup,
            args=(mode, base_dir, subject, zip_mode),
            daemon=True,
        ).start()

    # Стартира driver backup и връща резултата.
    def _run_driver_backup(self, mode: str, base_dir: Path, subject: str, zip_mode: str) -> None:
        try:
            backup_dir = create_backup_folder(base_dir)
            self.root.after(0, lambda: self._update_activation_progress(15, "Creating backup folder...", str(backup_dir)))

            result, log_path = export_drivers(backup_dir, mode)
            if result.stdout.strip():
                self.root.after(0, lambda text=result.stdout.strip(): self._append_activation_log(text))
            if result.stderr.strip():
                self.root.after(0, lambda text=result.stderr.strip(): self._append_activation_log(text))
            if result.returncode != 0:
                raise RuntimeError((result.stderr or result.stdout or "Driver export failed.").strip())

            self.root.after(0, lambda: self._update_activation_progress(50, "Creating driver list...", "pnputil /enum-drivers"))
            drivers_list_path = create_driver_list(backup_dir)

            self.root.after(0, lambda: self._update_activation_progress(70, "Creating restore script...", "RESTORE_DRIVERS.bat"))
            restore_script_path = create_restore_script(backup_dir)

            zip_path = None
            if zip_mode in {"keep", "delete"}:
                self.root.after(0, lambda: self._update_activation_progress(85, "Creating ZIP archive...", f"{backup_dir}.zip"))
                zip_path = compress_backup(backup_dir, delete_original=(zip_mode == "delete"))

            effective_backup_dir = backup_dir if backup_dir.exists() else Path(str(zip_path).removesuffix(".zip"))
            self.settings["last_driver_backup_dir"] = str(effective_backup_dir)
            self.settings["last_driver_backup_zip"] = str(zip_path) if zip_path else ""
            self.settings["last_driver_backup_log"] = str(log_path)
            self.settings["last_driver_list_path"] = str(drivers_list_path)
            self.settings["last_driver_restore_script"] = str(restore_script_path)
            save_settings(self.settings)

            details = [
                f"Backup folder: {effective_backup_dir}",
                f"Log: {log_path}",
                f"Driver list: {drivers_list_path}",
                f"Restore script: {restore_script_path}",
            ]
            if zip_path:
                details.append(f"ZIP archive: {zip_path}")
            self.root.after(0, lambda: self._finish_driver_workflow(subject, True, "\n".join(details)))
        except Exception as exc:
            self.root.after(0, lambda: self._finish_driver_workflow(subject, False, str(exc)))

    # Създава driver recovery usb и връща резултата към приложението.
    def _create_driver_recovery_usb(self) -> None:
        last_backup_dir = self._last_driver_backup_dir()
        if not last_backup_dir:
            messagebox.showerror("No Backup Found", "Run a driver backup first, then create the recovery USB.", parent=self.root)
            self.status_var.set("Recovery USB could not start because no backup was found.")
            return

        usb_drives = detect_removable_drives()
        if not usb_drives:
            messagebox.showerror("USB Not Found", "No removable USB drive was detected.", parent=self.root)
            self.status_var.set("Recovery USB could not start because no USB drive was detected.")
            return

        usb_root = usb_drives[0]
        confirmed = messagebox.askyesno(
            "Create Recovery USB",
            f"Use USB drive {usb_root} for DriverRecoveryBackup?",
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set("Recovery USB creation was canceled.")
            return

        self.status_var.set("Creating Recovery USB...")
        self._open_activation_window(
            title="Create Recovery USB",
            heading="Create Recovery USB",
            intro="The application is copying the last driver backup to USB and creating RESTORE_DRIVERS.bat.",
        )
        threading.Thread(target=self._run_create_recovery_usb, args=(last_backup_dir, usb_root), daemon=True).start()

    # Стартира create recovery usb и връща резултата.
    def _run_create_recovery_usb(self, backup_dir: Path, usb_root: Path) -> None:
        try:
            self.root.after(0, lambda: self._update_activation_progress(25, "Detecting USB destination...", str(usb_root)))
            recovery_dir, restore_script = create_recovery_usb(backup_dir, usb_root)
            self.settings["last_driver_recovery_usb"] = str(usb_root)
            save_settings(self.settings)
            details = f"USB: {usb_root}\nRecovery folder: {recovery_dir}\nRestore script: {restore_script}"
            self.root.after(0, lambda: self._finish_driver_workflow("Recovery USB", True, details))
        except Exception as exc:
            self.root.after(0, lambda: self._finish_driver_workflow("Recovery USB", False, str(exc)))

    # Помощна функция за generate driver pc report.
    def _generate_driver_pc_report(self) -> None:
        confirmed = messagebox.askyesno(
            "Generate PC Report",
            f"Create a new PC report on {desktop_path()} ?",
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set("PC report generation was canceled.")
            return

        self.status_var.set("Generating PC report...")
        self._open_activation_window(
            title="Generate PC Report",
            heading="Generate PC Report",
            intro="The application is collecting system information similar to the batch Speccy-like report.",
        )
        threading.Thread(target=self._run_generate_pc_report, daemon=True).start()

    # Стартира generate pc report и връща резултата.
    def _run_generate_pc_report(self) -> None:
        try:
            destination = create_backup_folder(desktop_path())
            self.root.after(0, lambda: self._update_activation_progress(20, "Creating report folder...", str(destination)))
            report_path = generate_pc_report(destination)
            self.settings["last_pc_report_path"] = str(report_path)
            save_settings(self.settings)
            self.root.after(0, lambda: self._finish_driver_workflow("PC Report", True, f"Report saved at:\n{report_path}"))
        except Exception as exc:
            self.root.after(0, lambda: self._finish_driver_workflow("PC Report", False, str(exc)))

    # Стартира driver backup advanced и връща резултата.
    def _run_driver_backup_advanced(self) -> None:
        base_dir = self._choose_driver_destination()
        if not base_dir:
            self.status_var.set("Advanced driver backup was canceled.")
            return

        backup_mode = self._choose_driver_backup_type()
        if not backup_mode:
            self.status_var.set("Advanced driver backup was canceled.")
            return

        zip_mode = self._choose_driver_zip_mode()
        if not zip_mode:
            self.status_var.set("Advanced driver backup was canceled.")
            return

        subject = f"Driver Backup Tool ({backup_mode.title()})"
        self._start_driver_backup(mode=backup_mode, base_dir=base_dir, subject=subject, zip_mode=zip_mode)

    # Възстановява drivers from last backup от подготвен backup.
    def _restore_drivers_from_last_backup(self) -> None:
        last_backup_dir = self._last_driver_backup_dir()
        if not last_backup_dir:
            messagebox.showerror("No Backup Found", "No saved backup folder was found.", parent=self.root)
            self.status_var.set("Driver restore could not start because no saved backup was found.")
            return

        confirmed = messagebox.askyesno(
            "Restore Drivers",
            f"Install drivers now from:\n\n{last_backup_dir}",
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set("Driver restore was canceled.")
            return

        self.status_var.set("Restoring drivers from last backup...")
        self._open_activation_window(
            title="Restore Drivers",
            heading="Restore Drivers",
            intro="The application is installing drivers from the last saved backup folder.",
        )
        threading.Thread(target=self._run_restore_drivers_from_last_backup, args=(last_backup_dir,), daemon=True).start()

    # Стартира restore drivers from last backup и връща резултата.
    def _run_restore_drivers_from_last_backup(self, backup_dir: Path) -> None:
        try:
            self.root.after(0, lambda: self._update_activation_progress(35, "Installing drivers from backup...", str(backup_dir)))
            result = restore_drivers_from_backup(backup_dir)
            details = "\n\n".join(part.strip() for part in (result.stdout, result.stderr) if part and part.strip()) or "Driver restore finished."
            if result.returncode != 0:
                raise RuntimeError(details)
            self.root.after(0, lambda: self._finish_driver_workflow("Driver Restore", True, details))
        except Exception as exc:
            self.root.after(0, lambda: self._finish_driver_workflow("Driver Restore", False, str(exc)))

    # Помощна функция за finish driver workflow.
    def _finish_driver_workflow(self, subject: str, success: bool, details: str) -> None:
        self._show_activation_result(success, details, subject)
        if self.current_menu == "driver_backup":
            self._render_cards()

    # Помощна функция за choose driver destination.
    def _choose_driver_destination(self) -> Path | None:
        dialog = tk.Toplevel(self.root)
        dialog.title("Driver Backup Destination")
        dialog.configure(bg="#0b1d0f")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        self._center_window(dialog, 470, 360)

        selected_path = tk.StringVar(value="")
        wrapper = tk.Frame(dialog, bg="#0b1d0f", padx=20, pady=18)
        wrapper.pack(fill="both", expand=True)

        tk.Label(wrapper, text="Choose Backup Destination", font=("Segoe UI Semibold", 16), fg="#d9ffe0", bg="#0b1d0f").pack(anchor="w", pady=(0, 6))
        tk.Label(wrapper, text="Matches the advanced batch tool: Desktop, USB, OneDrive or NAS path.", font=("Segoe UI", 10), fg="#9dc7a4", bg="#0b1d0f", wraplength=420, justify="left").pack(anchor="w", pady=(0, 14))

        # Помощна функция за choose.
        def choose(path: Path) -> None:
            selected_path.set(str(path))
            dialog.destroy()

        tk.Button(wrapper, text=f"Desktop\n{desktop_path()}", command=lambda: choose(desktop_path()), font=("Segoe UI Semibold", 10), bg="#174327", fg="#eefef1", activebackground="#236039", activeforeground="#ffffff", bd=0, padx=16, pady=10, cursor="hand2").pack(fill="x", pady=5)

        usb_drives = detect_removable_drives()
        if usb_drives:
            tk.Button(wrapper, text=f"USB Flash Drive\n{usb_drives[0]}", command=lambda: choose(usb_drives[0]), font=("Segoe UI Semibold", 10), bg="#174327", fg="#eefef1", activebackground="#236039", activeforeground="#ffffff", bd=0, padx=16, pady=10, cursor="hand2").pack(fill="x", pady=5)

        one_drive = onedrive_path()
        if one_drive:
            tk.Button(wrapper, text=f"OneDrive\n{one_drive}", command=lambda: choose(one_drive), font=("Segoe UI Semibold", 10), bg="#174327", fg="#eefef1", activebackground="#236039", activeforeground="#ffffff", bd=0, padx=16, pady=10, cursor="hand2").pack(fill="x", pady=5)

        # Помощна функция за choose nas.
        def choose_nas() -> None:
            nas_path = simpledialog.askstring("NAS Path", r"Enter NAS or network path, for example \\NAS\Backups", parent=dialog)
            if nas_path:
                selected_path.set(nas_path.strip())
                dialog.destroy()

        tk.Button(wrapper, text="NAS / Network Folder", command=choose_nas, font=("Segoe UI Semibold", 10), bg="#174327", fg="#eefef1", activebackground="#236039", activeforeground="#ffffff", bd=0, padx=16, pady=10, cursor="hand2").pack(fill="x", pady=5)
        tk.Button(wrapper, text="Cancel", command=dialog.destroy, font=("Segoe UI Semibold", 10), bg="#4c1c1c", fg="#fff4f4", activebackground="#7a1f1f", activeforeground="#ffffff", bd=0, padx=16, pady=10, cursor="hand2").pack(fill="x", pady=(12, 0))

        self.root.wait_window(dialog)
        selected = selected_path.get().strip()
        return Path(selected) if selected else None

    # Помощна функция за choose driver backup type.
    def _choose_driver_backup_type(self) -> str | None:
        dialog = tk.Toplevel(self.root)
        dialog.title("Driver Backup Type")
        dialog.configure(bg="#0b1d0f")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        self._center_window(dialog, 430, 260)

        selected = tk.StringVar(value="")
        wrapper = tk.Frame(dialog, bg="#0b1d0f", padx=20, pady=18)
        wrapper.pack(fill="both", expand=True)
        tk.Label(wrapper, text="Choose Backup Type", font=("Segoe UI Semibold", 16), fg="#d9ffe0", bg="#0b1d0f").pack(anchor="w", pady=(0, 10))

        tk.Button(wrapper, text="Full Backup (DISM)", command=lambda: (selected.set("full"), dialog.destroy()), font=("Segoe UI Semibold", 10), bg="#174327", fg="#eefef1", activebackground="#236039", activeforeground="#ffffff", bd=0, padx=16, pady=10, cursor="hand2").pack(fill="x", pady=5)
        tk.Button(wrapper, text="Clean Backup (PnPUtil)", command=lambda: (selected.set("clean"), dialog.destroy()), font=("Segoe UI Semibold", 10), bg="#174327", fg="#eefef1", activebackground="#236039", activeforeground="#ffffff", bd=0, padx=16, pady=10, cursor="hand2").pack(fill="x", pady=5)
        tk.Button(wrapper, text="Cancel", command=dialog.destroy, font=("Segoe UI Semibold", 10), bg="#4c1c1c", fg="#fff4f4", activebackground="#7a1f1f", activeforeground="#ffffff", bd=0, padx=16, pady=10, cursor="hand2").pack(fill="x", pady=(12, 0))

        self.root.wait_window(dialog)
        return selected.get() or None

    # Помощна функция за choose driver zip mode.
    def _choose_driver_zip_mode(self) -> str | None:
        dialog = tk.Toplevel(self.root)
        dialog.title("ZIP Compression")
        dialog.configure(bg="#0b1d0f")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        self._center_window(dialog, 460, 300)

        selected = tk.StringVar(value="")
        wrapper = tk.Frame(dialog, bg="#0b1d0f", padx=20, pady=18)
        wrapper.pack(fill="both", expand=True)
        tk.Label(wrapper, text="ZIP Compression", font=("Segoe UI Semibold", 16), fg="#d9ffe0", bg="#0b1d0f").pack(anchor="w", pady=(0, 10))

        tk.Button(wrapper, text="Create ZIP and keep original folder", command=lambda: (selected.set("keep"), dialog.destroy()), font=("Segoe UI Semibold", 10), bg="#174327", fg="#eefef1", activebackground="#236039", activeforeground="#ffffff", bd=0, padx=16, pady=10, cursor="hand2").pack(fill="x", pady=5)
        tk.Button(wrapper, text="Create ZIP and delete original folder", command=lambda: (selected.set("delete"), dialog.destroy()), font=("Segoe UI Semibold", 10), bg="#174327", fg="#eefef1", activebackground="#236039", activeforeground="#ffffff", bd=0, padx=16, pady=10, cursor="hand2").pack(fill="x", pady=5)
        tk.Button(wrapper, text="No ZIP, keep folder only", command=lambda: (selected.set("none"), dialog.destroy()), font=("Segoe UI Semibold", 10), bg="#174327", fg="#eefef1", activebackground="#236039", activeforeground="#ffffff", bd=0, padx=16, pady=10, cursor="hand2").pack(fill="x", pady=5)
        tk.Button(wrapper, text="Cancel", command=dialog.destroy, font=("Segoe UI Semibold", 10), bg="#4c1c1c", fg="#fff4f4", activebackground="#7a1f1f", activeforeground="#ffffff", bd=0, padx=16, pady=10, cursor="hand2").pack(fill="x", pady=(12, 0))

        self.root.wait_window(dialog)
        return selected.get() or None

    # Обработва събитието handle nexus admin action.
    def _handle_nexus_admin_action(self, action_id: str) -> None:
        status = self._nexus_admin_status()
        if not status.available:
            messagebox.showerror("Nexus Admin Unavailable", status.message, parent=self.root)
            self.status_var.set("Nexus Admin tools are not available on this system.")
            return

        if action_id == "nexus_list_users":
            self._run_nexus_background("List All Users", list_users)
            return
        if action_id == "nexus_change_password":
            username = simpledialog.askstring("Change Password", "Enter username:", parent=self.root)
            if not username:
                self.status_var.set("Password change was canceled.")
                return
            new_password = simpledialog.askstring("Change Password", f"Enter new password for {username}:", parent=self.root, show="*")
            if new_password is None or new_password == "":
                self.status_var.set("Password change was canceled.")
                return
            self._run_nexus_background("Change Password", lambda: change_password(username.strip(), new_password), subject=f"Password for {username.strip()}")
            return
        if action_id == "nexus_create_user":
            username = simpledialog.askstring("Create New User", "Enter username:", parent=self.root)
            username = username.strip() if username else ""
            if not username:
                self.status_var.set("User creation was canceled.")
                return
            wants_password = messagebox.askyesno("Create New User", f"Create user {username} with a password?", parent=self.root)
            password = None
            if wants_password:
                password = simpledialog.askstring("Create New User", f"Enter password for {username}:", parent=self.root, show="*")
                if password is None or password == "":
                    self.status_var.set("User creation was canceled.")
                    return
            make_admin = messagebox.askyesno("Create New User", f"Make {username} an Administrator?", parent=self.root)
            self._run_nexus_background(
                "Create New User",
                lambda: create_user(username, password, make_admin),
                subject=f"User {username}",
            )
            return
        if action_id == "nexus_delete_user":
            username = simpledialog.askstring("Delete User", "Enter the username to delete:", parent=self.root)
            username = username.strip() if username else ""
            if not username:
                self.status_var.set("User deletion was canceled.")
                return
            confirm_name = simpledialog.askstring(
                "Delete User",
                f'Type the username "{username}" again to confirm permanent deletion:',
                parent=self.root,
            )
            if (confirm_name or "").strip() != username:
                self.status_var.set("User deletion was canceled.")
                return
            self._run_nexus_background("Delete User", lambda: delete_user(username), subject=f"User {username}")
            return
        if action_id == "nexus_user_details":
            username = simpledialog.askstring("User Details", "Enter username:", parent=self.root)
            if not username:
                self.status_var.set("User details request was canceled.")
                return
            self._run_nexus_background("User Details", lambda: user_details(username.strip()), subject=username.strip())
            return
        if action_id == "nexus_toggle_admin":
            username = simpledialog.askstring("Administrator Rights", "Enter username:", parent=self.root)
            if not username:
                self.status_var.set("Administrator rights update was canceled.")
                return
            make_admin = messagebox.askyesno(
                "Administrator Rights",
                f"Choose Yes to add {username.strip()} to Administrators.\nChoose No to remove the user from Administrators.",
                parent=self.root,
            )
            self._run_nexus_background(
                "Administrator Rights",
                lambda: set_admin_rights(username.strip(), make_admin),
                subject=f"{username.strip()} admin rights",
            )

    # Стартира nexus background и връща резултата.
    def _run_nexus_background(self, title: str, runner: object, subject: str | None = None) -> None:
        self.status_var.set(f"Running {title}...")
        self._open_activation_window(
            title=title,
            heading=title,
            intro="The application is running the selected local account administration command.",
        )
        threading.Thread(
            target=self._run_nexus_command,
            args=(title, runner, subject or title),
            daemon=True,
        ).start()

    # Стартира nexus command и връща резултата.
    def _run_nexus_command(self, title: str, runner: object, subject: str) -> None:
        try:
            self.root.after(0, lambda: self._update_activation_progress(25, f"Starting {title}...", subject))
            result = runner()
            if isinstance(result, list):
                outputs: list[str] = []
                success = True
                for command_result in result:
                    text = "\n\n".join(part.strip() for part in (command_result.stdout, command_result.stderr) if part and part.strip())
                    if text:
                        outputs.append(text)
                        self.root.after(0, lambda line=text: self._append_activation_log(line))
                    if command_result.returncode != 0:
                        success = False
                details = "\n\n".join(outputs) or f"{title} finished."
                self.root.after(0, lambda: self._show_activation_result(success, details, subject))
                return

            details = "\n\n".join(part.strip() for part in (result.stdout, result.stderr) if part and part.strip()) or f"{title} finished."
            if details:
                self.root.after(0, lambda: self._append_activation_log(details))
            success = getattr(result, "returncode", 1) == 0
            self.root.after(0, lambda: self._show_activation_result(success, details, subject))
        except Exception as exc:
            self.root.after(0, lambda: self._show_activation_result(False, str(exc), subject))

    # Стартира инсталационната логика за office offline.
    def _install_office_offline(self, action_id: str) -> None:
        self._refresh_resource_panel()
        installer = get_office_offline_installer(action_id)
        office_info = self._office_install_info(action_id)
        missing_parts: list[str] = []
        if not installer.installers_root.exists():
            missing_parts.append(f"Installers folder not found: {installer.installers_root}")
        if not installer.setup_path.exists():
            missing_parts.append(f"setup.exe not found in {installer.setup_path.parent}")
        if not installer.config_path.exists():
            missing_parts.append(f"Config file not found: {installer.config_path.name}")

        if missing_parts:
            messagebox.showerror(
                "Office Installer Missing",
                "\n".join(missing_parts),
                parent=self.root,
            )
            self.status_var.set(f"{installer.label} could not start because installer files are missing.")
            return

        remove_existing = False
        if office_info.installed and office_info.uninstall_string:
            remove_existing = messagebox.askyesno(
                "Existing Office Found",
                (
                    f"Detected installed version:\n{office_info.display_name}\n\n"
                    "Remove the old version first and then install the selected one?"
                ),
                parent=self.root,
            )

        confirmed = messagebox.askyesno(
            "Start Office Installation",
            (
                f"Start offline installation for {installer.label} now?\n\n"
                f"Detected drive type: {self.launch_info['drive_type_label']}\n"
                f"Installers root: {installer.installers_root}"
            ),
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set(f"{installer.label} installation was canceled.")
            return

        self.status_var.set(f"Starting {installer.label}...")
        self._open_activation_window(
            title=f"{installer.label} Installation",
            heading=f"{installer.label} Setup",
            intro="Изпълнява се Office offline инсталация според локалните файлове в Installers папката.",
        )
        threading.Thread(
            target=self._run_office_offline_installation,
            args=(installer, remove_existing, office_info.display_name, office_info.uninstall_string),
            daemon=True,
        ).start()

    # Стартира инсталационната логика за office online.
    def _install_office_online(self, action_id: str) -> None:
        package = get_online_package(action_id)
        status = self._office_online_status(action_id)
        if not status.available:
            messagebox.showerror(
                "Online Package Not Available",
                status.message,
                parent=self.root,
            )
            self.status_var.set(f"{package.label} cannot start because the online package is not available.")
            return

        installed_now, installed_output, _ = self._office_online_install_state(action_id)
        remove_existing = False
        if installed_now:
            remove_existing = messagebox.askyesno(
                "Existing Package Found",
                (
                    f"Detected installed package for:\n{installed_output}\n\n"
                    "Remove the current version first and then install the new one?"
                ),
                parent=self.root,
            )

        confirmed = messagebox.askyesno(
            "Start Online Installation",
            f"Start online installation for {package.label} now?",
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set(f"{package.label} online installation was canceled.")
            return

        self.status_var.set(f"Starting online installation for {package.label}...")
        self._open_activation_window(
            title=f"{package.label} Online Installation",
            heading=f"{package.label} Online Setup",
            intro="Изпълнява се online инсталация чрез winget.",
        )
        threading.Thread(
            target=self._run_office_online_installation,
            args=(action_id, remove_existing),
            daemon=True,
        ).start()

    # Стартира инсталационната логика за local installer.
    def _install_local_installer(self, action_id: str) -> None:
        # Пуска локален installer от списъка с програми.
        spec = self._local_task_spec(action_id)
        if not spec:
            messagebox.showerror("Локален installer", "Липсва настройка за този installer.", parent=self.root)
            return
        local_file = self._find_resource_local_file(spec["resource_id"])
        if not local_file:
            messagebox.showerror(
                "Локален installer",
                f"Файлът за {spec['label']} не е намерен в папката Installers.",
                parent=self.root,
            )
            self.status_var.set(f"Липсва локален installer за {spec['label']}.")
            return

        installed_now, installed_text = self._task_install_state(spec)
        remove_existing = False
        if installed_now and self._task_supports_remove(spec):
            remove_existing = messagebox.askyesno(
                "Existing version found",
                f"Found:\n{installed_text}\n\nRemove it first and then start the new installer?",
                parent=self.root,
            )

        confirmed = messagebox.askyesno(
            "Local Installer",
            f"Start {spec['label']} from:\n\n{local_file}",
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set(f"Installation for {spec['label']} was canceled.")
            return

        self.status_var.set(f"Стартиране на {spec['label']}...")
        self._open_activation_window(
            title=spec["label"],
            heading=spec["label"],
            intro="Стартира се локален installer от папката Installers.",
        )
        threading.Thread(
            target=self._run_local_installer_installation,
            args=(action_id, local_file, remove_existing),
            daemon=True,
        ).start()

    # Стартира инсталационната логика за adobe reader.
    def _install_adobe_reader(self) -> None:
        self.adobe_reader_status_cache = None
        status = self._adobe_reader_status()
        winget_exe = find_winget_executable()
        latest = getattr(status, "latest_version", "") or "неизвестна"
        installed = getattr(status, "installed_version", "") or "не е открит"
        local_installer = getattr(status, "local_installer", None)
        local_version = getattr(status, "local_installer_version", "") or "неизвестна"

        details = (
            f"Актуална версия: {latest}\n"
            f"Инсталирана версия: {installed}\n"
            f"Локален installer: {local_installer or 'липсва'}\n"
            f"Версия на локалния installer: {local_version}\n\n"
            f"{getattr(status, 'message', '')}"
        )

        if not winget_exe:
            messagebox.showerror(
                "Adobe Reader",
                f"{details}\n\nWinget не е открит, затова не мога да изтегля актуалната версия автоматично.",
                parent=self.root,
            )
            self.status_var.set("Adobe Reader проверката приключи: winget липсва.")
            return

        installed_now, installed_output, uninstall_string = self._adobe_install_state()
        remove_existing = False
        if installed_now:
            remove_existing = messagebox.askyesno(
                "Existing Adobe Reader Found",
                (
                    f"Detected installed Adobe Reader version: {installed}\n\n"
                    "Remove the current version first and then install the latest one?"
                ),
                parent=self.root,
            )

        confirmed = messagebox.askyesno(
            "Adobe Reader",
            (
                f"{details}\n\n"
                "Install/update Adobe Reader to the current version through winget?"
            ),
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set("Adobe Reader installation was canceled.")
            return

        self.status_var.set("Starting Adobe Reader online installation...")
        self._open_activation_window(
            title="Adobe Reader",
            heading="Adobe Reader Online Setup",
            intro="Проверява се актуалната версия и се стартира инсталация чрез winget.",
        )
        threading.Thread(
            target=self._run_adobe_reader_installation,
            args=(winget_exe, remove_existing, installed_output, uninstall_string),
            daemon=True,
        ).start()

    # Стартира local installer installation и връща резултата.
    def _run_local_installer_installation(
        self,
        action_id: str,
        local_file: Path,
        remove_existing: bool = False,
    ) -> None:
        # Изпълнява локалния installer и показва статуса в прозореца за прогрес.
        spec = self._local_task_spec(action_id)
        if not spec:
            self.root.after(0, lambda: self._show_activation_result(False, "Липсва настройка за installer.", "Installer"))
            return
        output_lines: list[str] = []
        try:
            detect_mode = spec.get("detect_mode", "")
            detect_value = spec.get("detect_value", "")
            if remove_existing and detect_mode == "winget" and detect_value:
                installed_now, installed_output = self._is_winget_package_installed(detect_value)
                if installed_now:
                    self.root.after(
                        0,
                        lambda: self._update_activation_progress(
                            35,
                            f"Премахване на стара версия за {spec['label']}...",
                            installed_output or detect_value,
                        ),
                    )
                    winget_exe = find_winget_executable()
                    if winget_exe:
                        output_lines.append(self._run_winget_uninstall_command(winget_exe, detect_value, spec["label"]))

            command = [str(local_file)]
            silent_args = str(spec.get("silent_args", "")).strip()
            if silent_args:
                command.extend(part for part in silent_args.split(" ") if part)
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    60,
                    f"Стартиране на {spec['label']}...",
                    f"Р¤Р°Р№Р»: {local_file.name}",
                ),
            )
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
                cwd=str(local_file.parent),
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            output = self._collect_command_output(result)
            if output:
                output_lines.append(output)
                self.root.after(0, lambda text=output: self._append_activation_log(text))
            if result.returncode != 0:
                raise RuntimeError(output or f"{spec['label']} върна код {result.returncode}.")
            final_message = "\n\n".join(output_lines) or f"{spec['label']} завърши успешно."
            self.root.after(0, lambda: self._show_activation_result(True, final_message, spec["label"]))
        except Exception as exc:
            self.root.after(0, lambda: self._show_activation_result(False, str(exc), spec["label"]))

    # Стартира adobe reader installation и връща резултата.
    def _run_adobe_reader_installation(
        self,
        winget_exe: str,
        remove_existing: bool = False,
        installed_output: str = "",
        uninstall_string: str = "",
    ) -> None:
        command = [
            winget_exe,
            "install",
            "--id",
            ADOBE_READER_WINGET_ID,
            "--source",
            "winget",
            "--silent",
            "--disable-interactivity",
            "--accept-package-agreements",
            "--accept-source-agreements",
        ]
        output_lines: list[str] = []
        try:
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    20,
                    "Проверка на Adobe Reader пакета...",
                    f"Package ID: {ADOBE_READER_WINGET_ID}",
                ),
            )
            if remove_existing:
                self.root.after(
                    0,
                    lambda: self._update_activation_progress(
                        35,
                        "Открит е инсталиран Adobe Reader...",
                        installed_output or ADOBE_READER_WINGET_ID,
                    ),
                )
                winget_installed, _ = self._is_winget_package_installed(ADOBE_READER_WINGET_ID)
                if winget_installed:
                    removal_text = self._run_winget_uninstall_command(winget_exe, ADOBE_READER_WINGET_ID, "Adobe Reader")
                else:
                    removal_text = self._run_uninstall_string_command("Adobe Reader", uninstall_string)
                output_lines.append(removal_text)
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    60,
                    "Стартиране на Adobe Reader инсталация...",
                    f"Running: {' '.join(command)}",
                ),
            )
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if result.stdout.strip():
                output_lines.append(result.stdout.strip())
                self.root.after(0, lambda text=result.stdout.strip(): self._append_activation_log(text))
            if result.stderr.strip():
                output_lines.append(result.stderr.strip())
                self.root.after(0, lambda text=result.stderr.strip(): self._append_activation_log(text))
            if result.returncode != 0:
                raise RuntimeError("\n\n".join(output_lines) or f"Adobe Reader installer returned code {result.returncode}.")

            self.adobe_reader_status_cache = None
            final_message = "\n\n".join(output_lines) or "Adobe Reader беше инсталиран/обновен успешно."
            self.root.after(0, lambda: self._show_activation_result(True, final_message, "Adobe Reader"))
            self.root.after(0, self._render_cards)
        except Exception as exc:
            self.root.after(0, lambda: self._show_activation_result(False, str(exc), "Adobe Reader"))

    # Стартира office offline installation и връща резултата.
    def _run_office_offline_installation(
        self,
        installer: object,
        remove_existing: bool = False,
        existing_name: str = "",
        uninstall_string: str = "",
    ) -> None:
        output_lines: list[str] = []
        command = [
            str(installer.setup_path),
            "/configure",
            str(installer.config_path),
        ]
        try:
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    15,
                    "Проверка на Office installer файловете...",
                    f"Setup: {installer.setup_path}\nConfig: {installer.config_path}",
                ),
            )
            if remove_existing and uninstall_string:
                self.root.after(
                    0,
                    lambda: self._update_activation_progress(
                        35,
                        "Открита е стара Office версия...",
                        f"Подготвя се премахване на: {existing_name}",
                    ),
                )
                removal_text = self._run_office_uninstall_command(installer.action_id, existing_name, uninstall_string)
                output_lines.append(removal_text)
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    60,
                    f"Стартиране на {installer.label}...",
                    f"Running: {' '.join(command)}",
                ),
            )

            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                cwd=str(installer.setup_path.parent),
            )
            if result.stdout.strip():
                output_lines.append(result.stdout.strip())
                self.root.after(0, lambda text=result.stdout.strip(): self._append_activation_log(text))
            if result.stderr.strip():
                output_lines.append(result.stderr.strip())
                self.root.after(0, lambda text=result.stderr.strip(): self._append_activation_log(text))
            if result.returncode != 0:
                raise RuntimeError("\n\n".join(output_lines) or f"{installer.label} setup returned code {result.returncode}.")

            final_message = (
                "\n\n".join(output_lines)
                or f"{installer.label} installer finished successfully."
            )
            self.root.after(0, lambda: self._finish_office_installation(installer.action_id, installer.label, True, final_message))
        except Exception as exc:
            self.root.after(0, lambda: self._finish_office_installation(installer.action_id, installer.label, False, str(exc)))

    # Стартира office online installation и връща резултата.
    def _run_office_online_installation(
        self,
        action_id: str,
        remove_existing: bool = False,
    ) -> None:
        package = get_online_package(action_id)
        try:
            final_message = self._run_office_online_install_core(action_id, remove_existing=remove_existing)
            self.root.after(0, lambda: self._show_activation_result(True, final_message, package.label))
            self.root.after(0, self._render_cards)
        except Exception as exc:
            self.root.after(0, lambda: self._show_activation_result(False, str(exc), package.label))
        return
        output_lines: list[str] = []
        command = [
            winget_exe,
            "install",
            "--id",
            winget_id,
            "--source",
            "winget",
            "--silent",
            "--accept-package-agreements",
            "--accept-source-agreements",
        ]
        try:
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    15,
                    "Проверка на online пакета...",
                    f"Package ID: {winget_id}",
                ),
            )
            if remove_existing:
                self.root.after(
                    0,
                    lambda: self._update_activation_progress(
                        35,
                        "Открит е инсталиран пакет...",
                        installed_output or f"Package ID: {winget_id}",
                    ),
                )
                removal_text = self._run_winget_uninstall_command(winget_exe, winget_id, label)
                output_lines.append(removal_text)
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    60,
                    f"Стартиране на online инсталацията за {label}...",
                    f"Running: {' '.join(command)}",
                ),
            )
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if result.stdout.strip():
                output_lines.append(result.stdout.strip())
                self.root.after(0, lambda text=result.stdout.strip(): self._append_activation_log(text))
            if result.stderr.strip():
                output_lines.append(result.stderr.strip())
                self.root.after(0, lambda text=result.stderr.strip(): self._append_activation_log(text))
            if result.returncode != 0:
                raise RuntimeError("\n\n".join(output_lines) or f"{label} online installation returned code {result.returncode}.")

            final_message = "\n\n".join(output_lines) or f"{label} online installation started successfully."
            self.root.after(0, lambda: self._show_activation_result(True, final_message, label))
        except Exception as exc:
            self.root.after(0, lambda: self._show_activation_result(False, str(exc), label))

    # Проверява office activation status и връща резултат за интерфейса.
    def _check_office_activation_status(self) -> None:
        status = self._office_maintenance_status("office_check_activation_status")
        if not status.available:
            messagebox.showerror("OSPP Not Found", status.message, parent=self.root)
            self.status_var.set("Office activation status could not be checked because OSPP.VBS was not found.")
            return

        self.status_var.set("Checking Office activation status...")
        self._open_activation_window(
            title="Office Activation Status",
            heading="Office Activation Status",
            intro="The application is searching for OSPP.VBS and reading the activation status output.",
        )
        threading.Thread(target=self._run_office_activation_status, daemon=True).start()

    # Стартира office activation status и връща резултата.
    def _run_office_activation_status(self) -> None:
        ospp_vbs = find_ospp_vbs()
        if not ospp_vbs:
            self.root.after(0, lambda: self._show_activation_result(False, "OSPP.VBS was not found.", "Office"))
            return

        command = ["cscript", "//nologo", str(ospp_vbs), "/dstatus"]
        output_lines: list[str] = []
        try:
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    25,
                    "Searching for OSPP.VBS...",
                    f"Found: {ospp_vbs}",
                ),
            )
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    60,
                    "Reading Office activation status...",
                    f"Running: {' '.join(command)}",
                ),
            )
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if result.stdout.strip():
                output_lines.append(result.stdout.strip())
                self.root.after(0, lambda text=result.stdout.strip(): self._append_activation_log(text))
            if result.stderr.strip():
                output_lines.append(result.stderr.strip())
                self.root.after(0, lambda text=result.stderr.strip(): self._append_activation_log(text))
            if result.returncode != 0:
                raise RuntimeError("\n\n".join(output_lines) or f"OSPP status check returned code {result.returncode}.")

            final_message = "\n\n".join(output_lines) or "Office activation status was read successfully."
            self.root.after(0, lambda: self._show_activation_result(True, final_message, "Office"))
        except Exception as exc:
            self.root.after(0, lambda: self._show_activation_result(False, str(exc), "Office"))

    # Помощна функция за quick repair office.
    def _quick_repair_office(self) -> None:
        status = self._office_maintenance_status("office_quick_repair")
        if not status.available:
            messagebox.showerror("Repair Tool Not Found", status.message, parent=self.root)
            self.status_var.set("Office Quick Repair could not start because the repair tool was not found.")
            return

        confirmed = messagebox.askyesno(
            "Quick Repair Office",
            "Start the Office Click-to-Run repair workflow now?",
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set("Office Quick Repair was canceled.")
            return

        self.status_var.set("Starting Office Quick Repair...")
        self._open_activation_window(
            title="Office Quick Repair",
            heading="Office Quick Repair",
            intro="The application is starting the Office Click-to-Run repair workflow from the batch script.",
        )
        threading.Thread(target=self._run_office_quick_repair, daemon=True).start()

    # Стартира office quick repair и връща резултата.
    def _run_office_quick_repair(self) -> None:
        click_to_run = find_click_to_run_executable()
        if not click_to_run:
            self.root.after(0, lambda: self._show_activation_result(False, "OfficeClickToRun.exe was not found.", "Office"))
            return

        command = [
            str(click_to_run),
            "scenario=Repair",
            "platform=x64",
            "culture=en-us",
            "RepairType=FullRepair",
            "DisplayLevel=True",
        ]
        output_lines: list[str] = []
        try:
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    25,
                    "Checking Office repair tool...",
                    f"Found: {click_to_run}",
                ),
            )
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    55,
                    "Launching Office repair...",
                    f"Running: {' '.join(command)}",
                ),
            )
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if result.stdout.strip():
                output_lines.append(result.stdout.strip())
                self.root.after(0, lambda text=result.stdout.strip(): self._append_activation_log(text))
            if result.stderr.strip():
                output_lines.append(result.stderr.strip())
                self.root.after(0, lambda text=result.stderr.strip(): self._append_activation_log(text))
            if result.returncode != 0:
                raise RuntimeError("\n\n".join(output_lines) or f"Office repair returned code {result.returncode}.")

            final_message = "\n\n".join(output_lines) or "Office repair process launched successfully."
            self.root.after(0, lambda: self._show_activation_result(True, final_message, "Office"))
        except Exception as exc:
            self.root.after(0, lambda: self._show_activation_result(False, str(exc), "Office"))

    # Помощна функция за force uninstall all office.
    def _force_uninstall_all_office(self) -> None:
        status = self._office_maintenance_status("office_force_uninstall_all")
        if not status.available:
            messagebox.showerror("Winget Not Found", status.message, parent=self.root)
            self.status_var.set("Office cleanup could not start because winget is not available.")
            return

        confirm_text = simpledialog.askstring(
            "Force Uninstall Office",
            "This will try to remove all Office suites found on this PC.\n\nType CONFIRM to continue:",
            parent=self.root,
        )
        if (confirm_text or "").strip().upper() != "CONFIRM":
            self.status_var.set("Force uninstall was canceled.")
            return

        self.status_var.set("Starting Office cleanup...")
        self._open_activation_window(
            title="Force Uninstall Office",
            heading="Office Cleanup",
            intro="The application is running the same winget cleanup sequence defined in the batch script.",
        )
        threading.Thread(target=self._run_force_uninstall_all_office, daemon=True).start()

    # Стартира force uninstall all office и връща резултата.
    def _run_force_uninstall_all_office(self) -> None:
        winget_exe = find_winget_executable()
        if not winget_exe:
            self.root.after(0, lambda: self._show_activation_result(False, "Winget was not found.", "Office"))
            return

        output_lines: list[str] = []
        failures: list[str] = []
        total_steps = len(OFFICE_FORCE_UNINSTALL_IDS)
        try:
            for index, package_id in enumerate(OFFICE_FORCE_UNINSTALL_IDS, start=1):
                progress_value = 15 + int((index - 1) * 70 / max(1, total_steps))
                command = [winget_exe, "uninstall", "--id", package_id, "--silent"]
                self.root.after(
                    0,
                    lambda value=progress_value, pkg=package_id, cmd=command: self._update_activation_progress(
                        value,
                        f"Removing {pkg}...",
                        f"Running: {' '.join(cmd)}",
                    ),
                )
                result = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    check=False,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                command_output = "\n".join(part.strip() for part in (result.stdout, result.stderr) if part and part.strip())
                if command_output:
                    output_lines.append(f"[{package_id}]\n{command_output}")
                    self.root.after(0, lambda text=f"[{package_id}]\n{command_output}": self._append_activation_log(text))

                normalized_output = command_output.lower()
                if result.returncode == 0:
                    continue
                if "no installed package found" in normalized_output or "no package found matching input criteria" in normalized_output:
                    continue
                failures.append(package_id)

            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    100,
                    "Office cleanup finished.",
                    "The winget cleanup sequence has completed.",
                ),
            )
            final_message = "\n\n".join(output_lines) or "Office cleanup sequence finished."
            success = not failures
            if failures:
                final_message += "\n\nFailed packages:\n" + "\n".join(failures)
            self.root.after(0, lambda: self._show_activation_result(success, final_message, "Office"))
        except Exception as exc:
            self.root.after(0, lambda: self._show_activation_result(False, str(exc), "Office"))

    # Помощна функция за finish office installation.
    def _finish_office_installation(self, action_id: str, subject: str, success: bool, details: str) -> None:
        self.office_inventory_cache.pop(action_id, None)
        self._show_activation_result(success, details, subject)
        if self.current_menu == "office_install_center":
            self._render_cards()

    # Помощна функция за remove office installation.
    def _remove_office_installation(self, action_id: str) -> None:
        office_info = self._office_install_info(action_id)
        if not office_info.installed or not office_info.uninstall_string:
            self.status_var.set("No uninstall command was found for this Office version.")
            if self.current_menu == "office_install_center":
                self._render_cards()
            return

        confirmed = messagebox.askyesno(
            "Remove Office",
            f"Remove this Office installation?\n\n{office_info.display_name}",
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set("Office removal was canceled.")
            return

        self.status_var.set(f"Removing {office_info.display_name}...")
        self._open_activation_window(
            title="Remove Office",
            heading="Office Removal",
            intro="Изпълнява се деинсталация на намерената Office версия.",
        )
        threading.Thread(
            target=self._run_office_removal,
            args=(action_id, office_info.display_name, office_info.uninstall_string),
            daemon=True,
        ).start()

    # Стартира office removal и връща резултата.
    def _run_office_removal(self, action_id: str, display_name: str, uninstall_string: str) -> None:
        output_lines: list[str] = []
        try:
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    20,
                    "Подготовка на деинсталацията...",
                    f"Found uninstall command for {display_name}",
                ),
            )
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    55,
                    "Стартиране на премахването...",
                    f"Running: {uninstall_string}",
                ),
            )

            result = subprocess.run(
                ["cmd", "/c", uninstall_string],
                capture_output=True,
                text=True,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if result.stdout.strip():
                output_lines.append(result.stdout.strip())
                self.root.after(0, lambda text=result.stdout.strip(): self._append_activation_log(text))
            if result.stderr.strip():
                output_lines.append(result.stderr.strip())
                self.root.after(0, lambda text=result.stderr.strip(): self._append_activation_log(text))
            if result.returncode != 0:
                raise RuntimeError("\n\n".join(output_lines) or f"Uninstall command returned code {result.returncode}.")

            final_message = "\n\n".join(output_lines) or f"{display_name} removal finished."
            self.root.after(0, lambda: self._finish_office_removal(action_id, display_name, True, final_message))
        except Exception as exc:
            self.root.after(0, lambda: self._finish_office_removal(action_id, display_name, False, str(exc)))

    # Помощна функция за finish office removal.
    def _finish_office_removal(self, action_id: str, subject: str, success: bool, details: str) -> None:
        self.office_inventory_cache.pop(action_id, None)
        self._show_activation_result(success, details, subject)
        if self.current_menu == "office_install_center":
            self._render_cards()

    # Стартира office activation и връща резултата.
    def _run_office_activation(self, version_label: str, product_key: str) -> None:
        output_lines: list[str] = []
        try:
            commands = build_office_activation_commands(version_label, product_key)
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    10,
                    f"Preparing {version_label} activation environment...",
                    f"Starting {version_label} activation workflow.",
                ),
            )
            for progress_value, step_text, command in commands:
                masked_command = command[:-1] + ["[saved-key]"] if any("/inpkey:" in part for part in command) else command
                self.root.after(
                    0,
                    lambda value=progress_value, step=step_text, cmd=masked_command: self._update_activation_progress(
                        value,
                        step,
                        f"Running: {' '.join(cmd)}",
                    ),
                )
                result = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    check=False,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                if result.stdout.strip():
                    output_lines.append(result.stdout.strip())
                    self.root.after(0, lambda text=result.stdout.strip(): self._append_activation_log(text))
                if result.stderr.strip():
                    output_lines.append(result.stderr.strip())
                    self.root.after(0, lambda text=result.stderr.strip(): self._append_activation_log(text))
                if result.returncode != 0:
                    raise RuntimeError("\n\n".join(output_lines) or f"{version_label} activation command failed.")
        except Exception as exc:
            self.root.after(0, lambda: self._show_activation_result(False, str(exc), version_label))
            return

        final_output = "\n\n".join(output_lines) or f"{version_label} activation completed successfully."
        self.root.after(0, lambda: self._show_activation_result(True, final_output, version_label))

    # Стартира windows activation и връща резултата.
    def _run_windows_activation(self, version_label: str, product_key: str) -> None:
        slmgr_path = Path(os.environ.get("WINDIR", r"C:\Windows")) / "System32" / "slmgr.vbs"
        commands = [
            (
                45,
                "Installing product key...",
                ["cscript", "//nologo", str(slmgr_path), "/ipk", product_key],
            ),
            (
                90,
                "Requesting Microsoft activation...",
                ["cscript", "//nologo", str(slmgr_path), "/ato"],
            ),
        ]

        output_lines: list[str] = []
        try:
            self.root.after(0, lambda: self._update_activation_progress(10, "Preparing activation environment...", f"Starting {version_label} activation workflow."))
            for progress_value, step_text, command in commands:
                self.root.after(0, lambda value=progress_value, step=step_text, cmd=command: self._update_activation_progress(value, step, f"Running: {' '.join(cmd[:-1]) if '/ipk' in cmd else ' '.join(cmd)}"))
                result = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    check=False,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                if result.stdout.strip():
                    output_lines.append(result.stdout.strip())
                    self.root.after(0, lambda text=result.stdout.strip(): self._append_activation_log(text))
                if result.stderr.strip():
                    output_lines.append(result.stderr.strip())
                    self.root.after(0, lambda text=result.stderr.strip(): self._append_activation_log(text))
                if result.returncode != 0:
                    raise RuntimeError("\n\n".join(output_lines) or "Activation command failed.")
        except Exception as exc:
            self.root.after(0, lambda: self._show_activation_result(False, str(exc), version_label))
            return

        final_output = "\n\n".join(output_lines) or f"{version_label} activation completed successfully."
        self.root.after(0, lambda: self._show_activation_result(True, final_output, version_label))

    # Показва activation result в интерфейса.
    def _show_activation_result(self, success: bool, details: str, subject: str) -> None:
        if success:
            self.status_var.set(f"{subject} activation completed.")
            self._update_activation_progress(100, "Activation completed.", details, finished=True, success=True)
            return

        self.status_var.set(f"{subject} activation failed.")
        self._update_activation_progress(100, "Activation failed.", details, finished=True, success=False)

    # Отваря activation window или съответния прозорец.
    def _open_activation_window(self, title: str, heading: str, intro: str) -> None:
        if self.activation_window is not None and self.activation_window.winfo_exists():
            self.activation_window.destroy()

        self.activation_window = tk.Toplevel(self.root)
        self.activation_window.title(title)
        self.activation_window.geometry("560x360")
        self.activation_window.configure(bg="#0b1d0f")
        self.activation_window.resizable(False, False)
        self.activation_window.transient(self.root)

        wrapper = tk.Frame(self.activation_window, bg="#0b1d0f", padx=18, pady=18)
        wrapper.pack(fill="both", expand=True)

        tk.Label(
            wrapper,
            text=heading,
            font=("Segoe UI Semibold", 18),
            fg="#c9ffd0",
            bg="#0b1d0f",
        ).pack(anchor="w")

        tk.Label(
            wrapper,
            text=intro,
            font=("Segoe UI", 10),
            fg="#97c79d",
            bg="#0b1d0f",
        ).pack(anchor="w", pady=(4, 14))

        self.activation_status_var = tk.StringVar(value="Preparing activation environment...")
        tk.Label(
            wrapper,
            textvariable=self.activation_status_var,
            font=("Segoe UI Semibold", 11),
            fg="#e9ffec",
            bg="#0b1d0f",
        ).pack(anchor="w", pady=(0, 10))

        self.activation_progress_var = tk.IntVar(value=0)
        ttk.Progressbar(
            wrapper,
            orient="horizontal",
            length=500,
            mode="determinate",
            maximum=100,
            variable=self.activation_progress_var,
        ).pack(fill="x", pady=(0, 14))

        self.activation_log_widget = tk.Text(
            wrapper,
            height=10,
            bg="#08130a",
            fg="#d6f8da",
            insertbackground="#d6f8da",
            relief="flat",
            wrap="word",
            font=("Consolas", 9),
        )
        self.activation_log_widget.pack(fill="both", expand=True)
        self.activation_log_widget.insert("end", "Waiting for activation steps...\n")
        self.activation_log_widget.config(state="disabled")

        self.activation_close_button = tk.Button(
            wrapper,
            text="Close",
            command=self.activation_window.destroy,
            font=("Segoe UI Semibold", 10),
            bg="#174327",
            fg="#eefef1",
            activebackground="#236039",
            activeforeground="#ffffff",
            bd=0,
            padx=18,
            pady=8,
            state="disabled",
            cursor="hand2",
        )
        self.activation_close_button.pack(anchor="e", pady=(14, 0))

    # Помощна функция за append activation log.
    def _append_activation_log(self, text: str) -> None:
        if self.activation_log_widget is None or not self.activation_log_widget.winfo_exists():
            return
        self.activation_log_widget.config(state="normal")
        self.activation_log_widget.insert("end", f"{text}\n\n")
        self.activation_log_widget.see("end")
        self.activation_log_widget.config(state="disabled")

    # Обновява activation progress след промяна в състоянието.
    def _update_activation_progress(
        self,
        value: int,
        status_text: str,
        details: str,
        finished: bool = False,
        success: bool = False,
    ) -> None:
        if self.activation_window is None or not self.activation_window.winfo_exists():
            return
        if self.activation_progress_var is not None:
            self.activation_progress_var.set(value)
        if self.activation_status_var is not None:
            self.activation_status_var.set(status_text)
        self._append_activation_log(details)
        if finished and self.activation_close_button is not None:
            self.activation_close_button.config(state="normal", bg="#1d5a28" if success else "#7a1f1f")

    # Помощна функция за go back.
    def go_back(self) -> None:
        if not self.history:
            return
        target = self.history.pop()
        if target == "main":
            self.go_dashboard()
            self.status_var.set("Returned to Dashboard.")
            return
        self.render_menu(target)
        self.status_var.set(f"Returned to {MENU_TREE[target]['title']}.")

    # Помощна функция за go home.
    def go_home(self) -> None:
        self._stop_dashboard_info_scroll()
        if self.dashboard_render_job is not None:
            try:
                self.root.after_cancel(self.dashboard_render_job)
            except tk.TclError:
                pass
            self.dashboard_render_job = None
        self._show_dashboard_direct(reset_history=True)

    # Помощна функция за go dashboard.
    def go_dashboard(self) -> None:
        self.go_home()

    # Помощна функция за next page.
    def next_page(self) -> None:
        items = MENU_TREE[self.current_menu]["items"]
        page_size = MENU_PAGE_SIZE.get(self.current_menu, CARDS_PER_PAGE)
        total_pages = max(1, math.ceil(len(items) / page_size))
        if self.current_page < total_pages - 1:
            self.current_page += 1
            self._render_cards()

    # Помощна функция за previous page.
    def previous_page(self) -> None:
        if self.current_page > 0:
            self.current_page -= 1
            self._render_cards()


# Помощна функция за main.
def main() -> None:
    configure_windows_dpi_awareness()
    # Главна входна точка: иска admin права и после стартира UI.
    if not is_running_as_admin():
        started = relaunch_as_admin()
        if not started:
            temp_root = tk.Tk()
            temp_root.withdraw()
            messagebox.showerror(
                "Administrator Rights Required",
                "WinSys Guardian Advanced must run with administrator rights.",
                parent=temp_root,
            )
            temp_root.destroy()
        return

    root = tk.Tk()
    apply_app_icon(root)
    ProductLauncher(root)
    root.mainloop()


if __name__ == "__main__":
    main()
