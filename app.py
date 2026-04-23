from __future__ import annotations

import base64
import ctypes
import hashlib
import math
import os
import platform
import queue
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
import json
import webbrowser
from tkinter import messagebox, simpledialog, ttk
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
from office_installers import get_office_offline_installer
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
from office_online import check_online_package, find_winget_executable, get_online_package
from path_utils import get_runtime_storage_info
from resource_manager import (
    ResourceStatus,
    check_resource_status,
    download_resource,
    missing_resource_report,
)
from self_updater import launch_helper_and_exit, prepare_update_install
from system_health import HealthItem, collect_health_items
from update_checker import UpdateResult, check_for_updates


APP_TITLE = "WinSys Guardian Advanced"
WINDOW_SIZE = "930x630"
MAIN_WINDOW_SIZE = "1220x820"
MAIN_MIN_WIDTH = 1120
MAIN_MIN_HEIGHT = 760
MAIN_CARD_COLUMNS = 3
PROJECT_ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
SETTINGS_FILE = PROJECT_ROOT / "settings.json"
SECURE_STORE_FILE = PROJECT_ROOT / ".wga_secure_store.json"
VERSION_FILE = PROJECT_ROOT / "version.json"
APP_ICON_FILE = PROJECT_ROOT / "assets" / "wga-icon.ico"
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
MENU_PAGE_SIZE: dict[str, int] = {
    "activation": 4,
    "reset_onedrive": 4,
    "windows11_activation": 4,
    "office_activation": 4,
    "install_software": 4,
    "office_install_center": 4,
    "secret_install": 4,
    "office_center": 4,
    "language": 4,
    "driver_backup": 4,
    "nexus_admin": 4,
}

OFFICE_ACTION_IDS = [
    "office_2016_activation",
    "office_2019_activation",
    "office_2021_activation",
]


MENU_TREE = {
    "main": {
        "title": "Main Menu",
        "subtitle": "Central control hub for deployment, activation, language, recovery and admin tools.",
        "items": [
            {"label": "Activation Menu", "kind": "menu", "target": "activation"},
            {"label": "Add Desktop Icons", "kind": "action", "description": "Create standard support shortcuts."},
            {
                "label": "Reset OneDrive",
                "kind": "menu",
                "target": "reset_onedrive",
            },
            {"label": "Install Software", "kind": "menu", "target": "install_software"},
            {"label": "Language Menu", "kind": "menu", "target": "language"},
            {
                "label": "Driver Backup + PC Report",
                "kind": "menu",
                "target": "driver_backup",
            },
            {
                "label": "System Commander: Nexus Admin",
                "kind": "menu",
                "target": "nexus_admin",
                "description": "Local user management and administrator account tools.",
            },
            {"label": "Reset Console", "kind": "action", "description": "Refresh the current interface state."},
            {"label": "Exit", "kind": "exit", "description": "Close WinSys Guardian Advanced."},
        ],
    },
    "activation": {
        "title": "Activation Menu",
        "subtitle": "Windows and Office activation shortcuts.",
        "items": [
            {"label": "Activate Windows 10", "kind": "action"},
            {"label": "Activate Windows 11", "kind": "menu", "target": "windows11_activation"},
            {"label": "Office Activation Center", "kind": "menu", "target": "office_activation"},
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
                "description": "Run the Windows activation commands using the saved product key.",
            },
            {
                "label": "Save or Replace Product Key",
                "kind": "action",
                "action_id": "save_windows11_key",
                "description": "Store the currently approved Windows 11 key for later use.",
            },
            {
                "label": "Show Saved Product Key",
                "kind": "action",
                "action_id": "show_windows11_key",
                "description": "Display the key currently stored in this application.",
            },
            {
                "label": "Clear Saved Product Key",
                "kind": "action",
                "action_id": "clear_windows11_key",
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
                "description": "Store the Office product key used by your admin workflow.",
            },
            {
                "label": "Show Saved Office Key",
                "kind": "action",
                "action_id": "show_office_key",
                "description": "Display the Office key currently stored in this application.",
            },
            {
                "label": "Clear Saved Office Key",
                "kind": "action",
                "action_id": "clear_office_key",
                "description": "Remove the saved Office key if your organization replaces it.",
            },
            {
                "label": "Office 2016",
                "kind": "action",
                "action_id": "office_2016_activation",
                "description": "Run activation workflow for Office 2016 using the saved Office key.",
            },
            {
                "label": "Office 2019",
                "kind": "action",
                "action_id": "office_2019_activation",
                "description": "Run activation workflow for Office 2019 using the saved Office key.",
            },
            {
                "label": "Office 2021",
                "kind": "action",
                "action_id": "office_2021_activation",
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
            {"label": "Install Ninite", "kind": "action"},
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


def load_settings() -> dict[str, str]:
    if not SETTINGS_FILE.exists():
        return {}
    try:
        data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}
    return data if isinstance(data, dict) else {}


def save_settings(settings: dict[str, str]) -> None:
    SETTINGS_FILE.write_text(json.dumps(settings, indent=2), encoding="utf-8")


def load_version_info() -> dict[str, str]:
    defaults = {
        "version": "0.1.1",
        "version_info_url": "",
        "download_url": "",
        "notes": "",
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
    merged.update({key: str(value) for key, value in data.items() if value is not None})
    return merged


def _portable_secret_key() -> bytes:
    secret_seed = f"{APP_TITLE}|WGA-Portable-Store|{PROJECT_ROOT.name}"
    return hashlib.sha256(secret_seed.encode("utf-8")).digest()


def encrypt_for_current_user(text: str) -> str:
    source = text.encode("utf-8")
    key = _portable_secret_key()
    encrypted = bytes(byte ^ key[index % len(key)] for index, byte in enumerate(source))
    return base64.b64encode(encrypted).decode("ascii")


def decrypt_for_current_user(encoded_text: str) -> str:
    encrypted = base64.b64decode(encoded_text.encode("ascii"))
    key = _portable_secret_key()
    decrypted = bytes(byte ^ key[index % len(key)] for index, byte in enumerate(encrypted))
    return decrypted.decode("utf-8")


def hash_secret(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def ensure_hidden_file(path: Path) -> None:
    ctypes.windll.kernel32.SetFileAttributesW(str(path), FILE_ATTRIBUTE_HIDDEN)


def ensure_normal_file(path: Path) -> None:
    if path.exists():
        ctypes.windll.kernel32.SetFileAttributesW(str(path), FILE_ATTRIBUTE_NORMAL)


def get_drive_label(path: Path) -> str:
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


def get_launch_location_info() -> dict[str, str]:
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


def load_secure_store() -> dict[str, str]:
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


def save_secure_store(store: dict[str, str]) -> None:
    serialized = json.dumps(store, indent=2)
    encrypted = encrypt_for_current_user(serialized)
    payload = {"data": encrypted}
    ensure_normal_file(SECURE_STORE_FILE)
    SECURE_STORE_FILE.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    ensure_hidden_file(SECURE_STORE_FILE)


def is_running_as_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except OSError:
        return False


def apply_app_icon(root: tk.Tk | tk.Toplevel) -> None:
    if not APP_ICON_FILE.exists():
        return
    try:
        root.iconbitmap(str(APP_ICON_FILE))
    except tk.TclError:
        pass


def relaunch_as_admin() -> bool:
    script_path = str(Path(__file__).resolve())
    result = ctypes.windll.shell32.ShellExecuteW(
        None,
        "runas",
        sys.executable,
        f'"{script_path}"',
        None,
        1,
    )
    return result > 32
class SplashScreen:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        apply_app_icon(self.root)
        self.root.geometry(WINDOW_SIZE)
        self.root.configure(bg="black")
        self.root.resizable(False, False)

        self.progress_value = 0.0
        self.target_value = 0.0
        self.status_text = tk.StringVar(value="Initializing core modules...")
        self.message_queue: queue.Queue[tuple[str, float | str]] = queue.Queue()

        self.canvas = tk.Canvas(self.root, width=930, height=630, highlightthickness=0, bd=0)
        self.canvas.pack(fill="both", expand=True)

        self._draw_background()
        self._create_loader()
        self._start_boot_sequence()

    def _draw_background(self) -> None:
        width = 930
        height = 630
        band_count = 240

        self.canvas.create_rectangle(0, 0, width, height, fill="#000000", outline="")

        for index in range(band_count):
            x0 = index * width / band_count
            x1 = (index + 1) * width / band_count
            distance = abs((x0 + x1) / 2 - width / 2) / (width / 2)
            green = int(40 + (1 - distance) * 150)
            shade = max(0, min(255, green))
            color = f"#{0:02x}{shade:02x}{0:02x}"
            self.canvas.create_rectangle(x0, 0, x1, height, fill=color, outline=color)

        for index in range(160):
            x = (index * 37) % width
            offset = abs(x - width / 2) / (width / 2)
            line_width = 1 if index % 4 else 2
            intensity = max(55, int(255 - offset * 180))
            line_color = f"#{0:02x}{intensity:02x}{30:02x}"
            self.canvas.create_line(x, 40, x, height - 20, fill=line_color, width=line_width)

        for step in range(110, 0, -1):
            alpha_band = int(115 * (step / 110))
            radius_x = 175 + step * 2
            radius_y = 260 + step * 2
            left = width / 2 - radius_x / 2
            top = height / 2 - radius_y / 2
            right = width / 2 + radius_x / 2
            bottom = height / 2 + radius_y / 2
            color = f"#{0:02x}{alpha_band // 2:02x}{0:02x}"
            self.canvas.create_oval(left, top, right, bottom, fill=color, outline="")

    def _create_loader(self) -> None:
        self.canvas.create_text(465, 120, text=APP_TITLE, fill="#d9ffd9", font=("Segoe UI Semibold", 24))
        self.canvas.create_text(465, 160, text="System Startup Interface", fill="#7cf97c", font=("Segoe UI", 12))

        self.bar_left = 310
        self.bar_top = 340
        self.bar_width = 310
        self.bar_height = 30
        self.bar_radius = 12

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
            465,
            self.bar_top + self.bar_height / 2,
            text="0%",
            fill="#111111",
            font=("Segoe UI Semibold", 14),
        )
        self.status_label = self.canvas.create_text(
            465,
            395,
            text=self.status_text.get(),
            fill="#d4ffd4",
            font=("Segoe UI", 11),
        )

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

    def _start_boot_sequence(self) -> None:
        threading.Thread(target=self._run_startup_tasks, daemon=True).start()
        self._poll_queue()
        self._animate_progress()

    def _run_startup_tasks(self) -> None:
        tasks = [
            ("Loading configuration", 0.10, self._load_configuration),
            ("Collecting system profile", 0.28, self._collect_system_profile),
            ("Checking workspace files", 0.52, self._scan_workspace),
            ("Preparing security engines", 0.77, self._prepare_engines),
            ("Launching interface", 1.00, self._finalize_startup),
        ]
        for label, progress, action in tasks:
            self.message_queue.put(("status", label))
            action()
            self.message_queue.put(("progress", progress))
        self.message_queue.put(("done", "Ready"))

    def _load_configuration(self) -> None:
        config_file = PROJECT_ROOT / "tasks.json"
        if config_file.exists():
            config_file.read_text(encoding="utf-8")
        time.sleep(0.15)

    def _collect_system_profile(self) -> None:
        platform.platform()
        platform.processor()
        os.cpu_count()
        total = 0
        for number in range(1, 55_000):
            total += math.isqrt(number)
        _ = total

    def _scan_workspace(self) -> None:
        file_count = 0
        for path in PROJECT_ROOT.rglob("*"):
            if path.is_file():
                file_count += 1
            if file_count >= 500:
                break
        time.sleep(0.1)

    def _prepare_engines(self) -> None:
        total = 0
        for number in range(25_000):
            total += (number * number) % 97
        _ = total

    def _finalize_startup(self) -> None:
        time.sleep(0.25)

    def _poll_queue(self) -> None:
        try:
            while True:
                message_type, payload = self.message_queue.get_nowait()
                if message_type == "status":
                    self.status_text.set(str(payload))
                    self.canvas.itemconfig(self.status_label, text=self.status_text.get())
                elif message_type == "progress":
                    self.target_value = float(payload)
                elif message_type == "done":
                    self.status_text.set("System ready.")
                    self.canvas.itemconfig(self.status_label, text=self.status_text.get())
                    self.target_value = 1.0
                    self.root.after(700, self._show_dashboard)
        except queue.Empty:
            pass
        self.root.after(40, self._poll_queue)

    def _animate_progress(self) -> None:
        if self.progress_value < self.target_value:
            delta = max(0.004, (self.target_value - self.progress_value) * 0.18)
            self.progress_value = min(self.target_value, self.progress_value + delta)
            self._update_bar()
        self.root.after(16, self._animate_progress)

    def _update_bar(self) -> None:
        fill_width = max(4, self.bar_width * self.progress_value)
        points = self._rounded_rect_points(
            self.bar_left,
            self.bar_top,
            self.bar_left + fill_width,
            self.bar_top + self.bar_height,
            self.bar_radius,
        )
        self.canvas.coords(self.progress_fill, *points)
        self.canvas.itemconfig(self.progress_label, text=f"{int(self.progress_value * 100)}%")

    def _show_dashboard(self) -> None:
        self.canvas.destroy()
        self.root.geometry(MAIN_WINDOW_SIZE)
        self.root.minsize(MAIN_MIN_WIDTH, MAIN_MIN_HEIGHT)
        self.root.resizable(True, True)
        try:
            self.root.state("zoomed")
        except tk.TclError:
            pass
        MainMenuUI(self.root)


class MainMenuUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.configure(bg="#08130a")
        self.settings = load_settings()
        self.secure_store = load_secure_store()
        self.launch_info = get_launch_location_info()
        self.resource_status: ResourceStatus = check_resource_status(PROJECT_ROOT)
        self.version_info = load_version_info()
        self.activation_window: tk.Toplevel | None = None
        self.activation_status_var: tk.StringVar | None = None
        self.activation_progress_var: tk.IntVar | None = None
        self.activation_log_widget: tk.Text | None = None
        self.activation_close_button: tk.Button | None = None
        self.health_rows: list[tuple[tk.Label, tk.Label, tk.Label]] = []
        self.health_canvas: tk.Canvas | None = None
        self.health_scrollbar: ttk.Scrollbar | None = None
        self.health_inner_frame: tk.Frame | None = None
        self.health_scroll_position = 0.0
        self.health_scroll_job: str | None = None
        self.update_result: UpdateResult | None = None
        self.update_download_url = ""
        self.update_package_url = ""
        self.update_installing = False
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

        self.container = tk.Frame(self.root, bg="#08130a")
        self.container.pack(fill="both", expand=True)

        self.header = tk.Frame(self.container, bg="#0b1d0f", height=90)
        self.header.pack(fill="x")
        self.header.pack_propagate(False)

        self.title_label = tk.Label(
            self.header,
            text=APP_TITLE,
            font=("Segoe UI Semibold", 22),
            fg="#8cff95",
            bg="#0b1d0f",
        )
        self.title_label.pack(anchor="w", padx=24, pady=(14, 0))

        self.header_exit_button = tk.Button(
            self.header,
            text="Изход",
            command=self.root.destroy,
            font=("Segoe UI Semibold", 10),
            bg="#7a1f1f",
            fg="#fff4f4",
            activebackground="#a32d2d",
            activeforeground="#ffffff",
            bd=0,
            padx=18,
            pady=8,
            width=10,
            cursor="hand2",
        )
        self.header_exit_button.place(relx=1.0, x=-24, y=22, anchor="ne")

        self.header_home_button = tk.Button(
            self.header,
            text="Главно меню",
            command=self.go_home,
            font=("Segoe UI Semibold", 10),
            bg="#174327",
            fg="#eefef1",
            activebackground="#236039",
            activeforeground="#ffffff",
            bd=0,
            padx=18,
            pady=8,
            width=12,
            cursor="hand2",
        )
        self.header_home_button.place(relx=1.0, x=-160, y=22, anchor="ne")

        self.subtitle_label = tk.Label(
            self.header,
            text="",
            font=("Segoe UI", 10),
            fg="#b9e8be",
            bg="#0b1d0f",
        )
        self.subtitle_label.pack(anchor="w", padx=24)

        self.version_chip = tk.Label(
            self.header,
            text=f"v{self.version_info['version']}",
            font=("Segoe UI Semibold", 9),
            fg="#d9ffe0",
            bg="#174327",
            padx=10,
            pady=4,
        )
        self.version_chip.place(x=540, y=26)

        self.content = tk.Frame(self.container, bg="#08130a")
        self.content.pack(fill="both", expand=True, padx=20, pady=18)

        self.left_panel = tk.Frame(self.content, bg="#0e1f11", width=320, bd=0, highlightthickness=1, highlightbackground="#174c1e")
        self.left_panel.pack(side="left", fill="y")
        self.left_panel.pack_propagate(False)

        self.menu_title = tk.Label(
            self.left_panel,
            text="Навигация",
            font=("Segoe UI Semibold", 15),
            fg="#b8ffbc",
            bg="#0e1f11",
        )
        self.menu_title.pack(anchor="w", padx=16, pady=(16, 8))

        self.menu_path = tk.Label(
            self.left_panel,
            text="Главно меню",
            justify="left",
            wraplength=280,
            font=("Segoe UI", 10),
            fg="#95c49b",
            bg="#0e1f11",
        )
        self.menu_path.pack(anchor="w", padx=16)

        self.system_info = tk.Label(
            self.left_panel,
            text=self._build_system_summary(),
            justify="left",
            wraplength=280,
            font=("Consolas", 9),
            fg="#d5f8d7",
            bg="#0e1f11",
        )
        self.system_info.pack(anchor="w", padx=16, pady=(20, 14))

        self.hint_label = tk.Label(
            self.left_panel,
            text="Избери карта отдясно, за да отвориш секция или да стартираш подготвено действие.",
            wraplength=280,
            justify="left",
            font=("Segoe UI", 9),
            fg="#8ca391",
            bg="#0e1f11",
        )
        self.hint_label.pack(anchor="w", padx=16)

        self.health_title = tk.Label(
            self.left_panel,
            text="Състояние на системата",
            font=("Segoe UI Semibold", 14),
            fg="#b8ffbc",
            bg="#0e1f11",
        )
        self.health_title.pack(anchor="w", padx=16, pady=(18, 8))

        self.health_frame = tk.Frame(
            self.left_panel,
            bg="#112716",
            bd=0,
            highlightthickness=1,
            highlightbackground="#1b5c25",
        )
        self.health_frame.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self.health_loading_label = tk.Label(
            self.health_frame,
            text="Loading hardware diagnostics...",
            font=("Segoe UI", 10),
            fg="#cdeed0",
            bg="#112716",
            justify="left",
            wraplength=260,
        )
        self.health_loading_label.pack(anchor="w", padx=12, pady=12)

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
            fg="#d5f8d7",
            bg="#0b1d0f",
            padx=18,
        )
        self.status_bar.pack(fill="x", side="bottom")

        self.right_panel = tk.Frame(self.content, bg="#08130a")
        self.right_panel.pack(side="left", fill="both", expand=True, padx=(18, 0))

        self.card_title = tk.Label(
            self.right_panel,
            text="",
            font=("Segoe UI Semibold", 19),
            fg="#e6ffee",
            bg="#08130a",
        )
        self.card_title.pack(anchor="w")

        self.card_subtitle = tk.Label(
            self.right_panel,
            text="",
            font=("Segoe UI", 10),
            fg="#9bc39e",
            bg="#08130a",
            wraplength=630,
            justify="left",
        )
        self.card_subtitle.pack(anchor="w", pady=(4, 12))

        self.update_banner = tk.Frame(
            self.right_panel,
            bg="#153042",
            bd=0,
            highlightthickness=1,
            highlightbackground="#2a5975",
        )
        self.update_banner.pack(fill="x", pady=(0, 12))

        self.update_icon_label = tk.Label(
            self.update_banner,
            text="i",
            font=("Segoe UI Semibold", 16),
            fg="#c7ecff",
            bg="#153042",
            width=2,
        )
        self.update_icon_label.pack(side="left", padx=(14, 8), pady=10)

        self.update_message_var = tk.StringVar(
            value=f"Проверка за актуализации за v{self.version_info['version']}..."
        )
        self.update_message_label = tk.Label(
            self.update_banner,
            textvariable=self.update_message_var,
            font=("Segoe UI", 10),
            fg="#d7f1ff",
            bg="#153042",
            justify="left",
            anchor="w",
        )
        self.update_message_label.pack(side="left", fill="x", expand=True, pady=10)

        self.update_action_button = tk.Button(
            self.update_banner,
            text="Отвори",
            command=self._open_update_download,
            font=("Segoe UI Semibold", 9),
            bg="#2b607a",
            fg="#f3fbff",
            activebackground="#3a7f9f",
            activeforeground="#ffffff",
            bd=0,
            padx=14,
            pady=7,
            state="disabled",
            cursor="hand2",
        )
        self.update_action_button.pack(side="right", padx=12, pady=8)

        self.resource_frame = tk.Frame(
            self.right_panel,
            bg="#112716",
            bd=0,
            highlightthickness=1,
            highlightbackground="#1b5c25",
        )
        self.resource_frame.pack(fill="x", pady=(0, 12))

        self.resource_title = tk.Label(
            self.resource_frame,
            text="\u0418\u043d\u0441\u0442\u0430\u043b\u0430\u0446\u0438\u043e\u043d\u043d\u0438 \u0440\u0435\u0441\u0443\u0440\u0441\u0438",
            font=("Segoe UI Semibold", 11),
            fg="#b8ffbc",
            bg="#112716",
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
            bg="#112716",
        )
        self.resource_status_label.pack(side="left", fill="x", expand=True, pady=10)

        self.resource_download_button = tk.Button(
            self.resource_frame,
            text="\u0418\u0437\u0442\u0435\u0433\u043b\u0438",
            command=self._download_missing_resources,
            font=("Segoe UI Semibold", 9),
            bg="#7d6a2d",
            fg="#fff7d6",
            activebackground="#9a8337",
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
            bg="#174327",
            fg="#eefef1",
            activebackground="#236039",
            activeforeground="#ffffff",
            bd=0,
            padx=12,
            pady=7,
            cursor="hand2",
        )
        self.resource_details_button.pack(side="right", pady=10)
        self._refresh_resource_panel()

        self.nav_frame = tk.Frame(
            self.right_panel,
            bg="#08130a",
            bd=0,
            highlightthickness=1,
            highlightbackground="#1f5928",
        )
        self.nav_frame.pack(fill="x", side="bottom", pady=(10, 0))

        self.page_label = tk.Label(
            self.nav_frame,
            text="Page 1 / 1",
            font=("Segoe UI", 10),
            fg="#b9e8be",
            bg="#08130a",
        )
        self.page_label.pack(side="left")

        self.controls_frame = tk.Frame(self.nav_frame, bg="#08130a")
        self.controls_frame.pack(side="right")

        self.prev_button = self._make_nav_button(self.controls_frame, "\u041d\u0430\u0437\u0430\u0434", self.previous_page)
        self.prev_button.pack(side="left", padx=(0, 6))
        self.next_button = self._make_nav_button(self.controls_frame, "\u041d\u0430\u043f\u0440\u0435\u0434", self.next_page)
        self.next_button.pack(side="left", padx=(0, 6))
        self.back_button = self._make_nav_button(self.controls_frame, "\u041d\u0430\u0437\u0430\u0434", self.go_back, accent="#17361f")
        self.back_button.pack(side="left", padx=(0, 6))
        self.home_button = self._make_nav_button(self.controls_frame, "\u041d\u0430\u0447\u0430\u043b\u043e", self.go_home, accent="#17361f")
        self.home_button.pack(side="left", padx=(0, 6))
        self.exit_button = self._make_nav_button(self.controls_frame, "\u0418\u0437\u0445\u043e\u0434", self.root.destroy, accent="#7a1f1f")
        self.exit_button.pack(side="left")

        self.cards_frame = tk.Frame(self.right_panel, bg="#08130a")
        self.cards_frame.pack(fill="both", expand=True)

        self.language_status_panel = tk.Frame(
            self.cards_frame,
            bg="#0e1f11",
            width=280,
            bd=0,
            highlightthickness=1,
            highlightbackground="#1b5c25",
        )
        self.language_status_panel.grid_propagate(False)

        self.language_status_title = tk.Label(
            self.language_status_panel,
            text="Език и клавиатури",
            font=("Segoe UI Semibold", 13),
            fg="#b8ffbc",
            bg="#0e1f11",
        )
        self.language_status_title.pack(anchor="w", padx=14, pady=(14, 6))

        self.language_status_frame = tk.Frame(
            self.language_status_panel,
            bg="#112716",
            bd=0,
            highlightthickness=1,
            highlightbackground="#1b5c25",
        )
        self.language_status_frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        self.language_status_label = tk.Label(
            self.language_status_frame,
            textvariable=self.language_status_var,
            justify="left",
            wraplength=230,
            font=("Segoe UI", 9),
            fg="#d5f8d7",
            bg="#112716",
        )
        self.language_status_label.pack(anchor="w", fill="x", padx=12, pady=10)

        self.render_menu("main", reset_history=True)
        self._load_language_status_async()
        self._load_system_health_async()
        self._check_updates_async()

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

    def _resource_status_color(self) -> str:
        if not self.resource_status.configured:
            return "#f9e6a8"
        if self.resource_status.complete:
            return "#9aff9f"
        if self.resource_status.downloadable_missing:
            return "#ffe08a"
        return "#ffb0a8"

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

    def _refresh_resource_panel(self) -> None:
        self.launch_info = get_launch_location_info()
        self.resource_status = check_resource_status(PROJECT_ROOT)
        if hasattr(self, "system_info"):
            self.system_info.config(text=self._build_system_summary())
        self.resource_status_label.config(
            text=self._build_resource_summary(),
            fg=self._resource_status_color(),
        )
        can_download = self.resource_status.missing > 0
        self.resource_download_button.config(
            state="normal" if can_download else "disabled",
            bg="#7d6a2d" if can_download else "#384039",
            fg="#fff7d6" if can_download else "#9aa69c",
        )

    def _show_resource_details(self) -> None:
        self._refresh_resource_panel()
        messagebox.showinfo(
            "\u0418\u043d\u0441\u0442\u0430\u043b\u0430\u0446\u0438\u043e\u043d\u043d\u0438 \u0440\u0435\u0441\u0443\u0440\u0441\u0438",
            missing_resource_report(self.resource_status),
            parent=self.root,
        )

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
        self.status_var.set("Изтегляне на липсващи инсталационни ресурси...")
        threading.Thread(target=self._run_resource_downloads, args=(missing_downloads,), daemon=True).start()

    def _run_resource_downloads(self, items: list[object]) -> None:
        errors: list[str] = []

        def progress(downloaded: int, total: int, name: str) -> None:
            percent = int((downloaded / total) * 100) if total else 0
            self.root.after(0, lambda: self.status_var.set(f"Изтегляне: {name} - {percent}%"))

        for item in items:
            try:
                download_resource(PROJECT_ROOT, item, progress)
            except Exception as exc:
                errors.append(f"{item.name}: {exc}")

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

    def _load_system_health_async(self) -> None:
        threading.Thread(target=self._collect_and_render_system_health, daemon=True).start()

    def _load_language_status_async(self) -> None:
        threading.Thread(target=self._collect_and_render_language_status, daemon=True).start()

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

    def _build_language_status_summary(self, status: LanguageStatus) -> str:
        def marker(value: bool) -> str:
            return "[OK]" if value else "[--]"

        return (
            f"{marker(status.has_bulgarian)} bg-BG в списъка\n"
            f"{marker(status.has_language_pack)} Български езиков пакет\n"
            f"{marker(status.has_bds)} БДС клавиатура\n"
            f"{marker(status.has_phonetic)} Фонетична стандартна\n"
            f"{marker(status.has_traditional)} Фонетична традиционна"
        )

    def _apply_language_status_summary(self, text: str, color: str) -> None:
        self.language_status_var.set(text)
        self.language_status_label.config(fg=color)

    def _collect_and_render_system_health(self) -> None:
        try:
            items = collect_health_items()
        except Exception as exc:
            items = [HealthItem(label="Health:", value=f"Diagnostics failed: {exc}", ok=False)]
        self.root.after(0, lambda: self._render_system_health(items))

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

    def _on_health_mousewheel(self, event: tk.Event) -> None:
        if self.health_canvas is None:
            return
        self.health_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

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

    def _check_updates_async(self) -> None:
        threading.Thread(target=self._perform_update_check, daemon=True).start()

    def _perform_update_check(self) -> None:
        result = check_for_updates(
            self.version_info["version"],
            self.version_info.get("version_info_url", ""),
        )
        self.root.after(0, lambda: self._apply_update_result(result))

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

    def _restart_command(self) -> list[str]:
        if getattr(sys, "frozen", False):
            return [sys.executable]
        return [sys.executable, str(Path(__file__).resolve())]

    def _install_update_package(self, package_url: str) -> None:
        if self.update_installing:
            return
        if not messagebox.askyesno(
            "Инсталиране на актуализация",
            "Приложението ще изтегли актуализацията, ще подмени файловете, след това ще се затвори и ще се отвори отново. Да продължим ли?",
            parent=self.root,
        ):
            return

        self.update_installing = True
        progress_window = tk.Toplevel(self.root)
        progress_window.title("WGA актуализация")
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

        def update_progress(value: int) -> None:
            self.root.after(0, lambda: progress_var.set(value))

        def worker() -> None:
            try:
                self.root.after(0, lambda: status_var.set("Сваляне на update пакета..."))
                helper_path = prepare_update_install(
                    package_url=package_url,
                    target_root=PROJECT_ROOT,
                    restart_command=self._restart_command(),
                    progress_callback=update_progress,
                )

                def finish() -> None:
                    status_var.set("Готово. Приложението ще се рестартира...")
                    progress_var.set(100)
                    launch_helper_and_exit(helper_path)
                    self.root.after(500, self.root.destroy)

                self.root.after(0, finish)
            except Exception as exc:
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

    def _make_nav_button(
        self,
        parent: tk.Widget,
        text: str,
        command: object,
        accent: str = "#113b18",
    ) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            font=("Segoe UI Semibold", 10),
            bg=accent,
            fg="#e6ffee",
            activebackground="#1d5a28",
            activeforeground="#ffffff",
            bd=0,
            padx=18,
            pady=10,
            width=NAV_BUTTON_WIDTH,
            cursor="hand2",
        )

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
            font=("Segoe UI Semibold", 10),
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground="#ffffff",
            bd=0,
            borderwidth=0,
            relief="flat",
            overrelief="flat",
            highlightthickness=0,
            takefocus=0,
            padx=16,
            pady=8,
            width=CARD_BUTTON_WIDTH,
            height=CARD_BUTTON_HEIGHT,
            wraplength=260,
            justify="center",
            cursor=cursor,
            state=state,
        )

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
        self.header_home_button.config(state="disabled" if menu_key == "main" else "normal")
        self._toggle_language_status_panel(menu_key == "language")
        self._render_cards()

    def _toggle_language_status_panel(self, visible: bool) -> None:
        if not visible and self.language_status_panel.winfo_ismapped():
            self.language_status_panel.grid_forget()

    def _build_path(self) -> str:
        trail = [MENU_TREE[key]["title"] for key in self.history + [self.current_menu]]
        return " > ".join(trail)

    def _render_cards(self) -> None:
        for widget in self.cards_frame.winfo_children():
            if widget is self.language_status_panel:
                widget.grid_forget()
                continue
            widget.destroy()
        for index in range(12):
            self.cards_frame.rowconfigure(index, weight=0, minsize=0)
            self.cards_frame.columnconfigure(index, weight=0, minsize=0)

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
        for row in range(row_count):
            self.cards_frame.rowconfigure(row, weight=1, minsize=CARD_MIN_HEIGHT)

        for index, item in enumerate(page_items):
            row = index // card_columns
            column = index % card_columns
            card = self._build_card(self.cards_frame, item)
            card.grid(row=row, column=column, sticky="nsew", padx=8, pady=8)

        if self.current_menu == "language":
            panel_column = card_columns
            self.cards_frame.columnconfigure(panel_column, weight=0, minsize=280)
            self.language_status_panel.grid(
                row=0,
                column=panel_column,
                rowspan=row_count,
                sticky="nsew",
                padx=(12, 8),
                pady=8,
            )
        elif self.language_status_panel.winfo_ismapped():
            self.language_status_panel.grid_forget()

        self.page_label.config(text=f"Page {self.current_page + 1} / {total_pages}")
        self.prev_button.config(state="normal" if self.current_page > 0 else "disabled")
        self.next_button.config(state="normal" if self.current_page < total_pages - 1 else "disabled")
        self.back_button.config(state="normal" if self.history else "disabled")

    def _card_columns(self) -> int:
        return MAIN_CARD_COLUMNS if self.current_menu == "main" else 2

    def _build_card(self, parent: tk.Widget, item: dict[str, str]) -> tk.Frame:
        accent = self._card_accent(item)
        card_bg = "#122d19" if item["kind"] == "menu" else "#102515"
        border_color = "#2d7f4a" if item["kind"] == "menu" else "#1f5928"
        card = tk.Frame(
            parent,
            bg=card_bg,
            bd=0,
            highlightthickness=1,
            highlightbackground=border_color,
            height=CARD_MIN_HEIGHT,
        )
        card.grid_propagate(False)

        top = tk.Frame(card, bg=card_bg)
        top.pack(fill="x", padx=14, pady=(12, 8))

        dot = tk.Canvas(top, width=16, height=16, bg=card_bg, highlightthickness=0)
        dot.create_oval(2, 2, 14, 14, fill=accent, outline="")
        dot.pack(side="left")

        title = tk.Label(
            top,
            text=item["label"],
            font=("Segoe UI Semibold", 12),
            fg="#edffef",
            bg=card_bg,
            anchor="w",
            wraplength=320,
            justify="left",
        )
        title.pack(side="left", padx=(8, 0), fill="x", expand=True)

        description = self._item_description(item)
        desc_label = tk.Label(
            card,
            text=description,
            font=("Segoe UI", 9),
            fg="#96b79a",
            bg=card_bg,
            wraplength=320,
            justify="left",
            anchor="nw",
        )
        desc_label.pack(fill="x", padx=14, pady=(0, 8))

        spacer = tk.Frame(card, bg=card_bg)
        spacer.pack(fill="both", expand=True)

        has_remove_button = False
        if self._is_office_install_item(item):
            office_info = self._office_install_info(item["action_id"])
            has_remove_button = bool(office_info.installed and office_info.uninstall_string)

        action_area_height = CARD_ACTION_DOUBLE_HEIGHT if has_remove_button else CARD_ACTION_HEIGHT
        action_area = tk.Frame(card, bg=card_bg, height=action_area_height)
        action_area.pack(fill="x", padx=14, pady=(16, 16), side="bottom")
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
            width=CARD_BUTTON_PIXEL_WIDTH,
            height=CARD_BUTTON_PIXEL_HEIGHT,
        )

        if self._is_office_install_item(item):
            if has_remove_button:
                remove_button = self._make_card_button(
                    action_area,
                    text="\u041f\u0440\u0435\u043c\u0430\u0445\u043d\u0438",
                    command=lambda selected=item: self._remove_office_installation(selected["action_id"]),
                    bg="#9a2f2f",
                    fg="#fff6f6",
                    active_bg="#c24040",
                    cursor="hand2",
                )
                remove_button.place(
                    relx=0.5,
                    y=CARD_BUTTON_PIXEL_HEIGHT + 8,
                    anchor="n",
                    width=CARD_BUTTON_PIXEL_WIDTH,
                    height=CARD_BUTTON_PIXEL_HEIGHT,
                )

        return card

    def _is_office_install_item(self, item: dict[str, str]) -> bool:
        action_id = item.get("action_id", "")
        return action_id.startswith("install_office_") and action_id.endswith("_offline")

    def _is_office_online_item(self, item: dict[str, str]) -> bool:
        return item.get("action_id", "").startswith("online_")

    def _is_office_maintenance_item(self, item: dict[str, str]) -> bool:
        return item.get("action_id", "") in {
            "office_check_activation_status",
            "office_quick_repair",
            "office_force_uninstall_all",
        }

    def _is_language_item(self, item: dict[str, str]) -> bool:
        return item.get("action_id", "") in {
            "language_refresh",
            "toggle_bulgarian_bds",
            "toggle_bulgarian_phonetic",
            "toggle_bulgarian_traditional",
            "toggle_bulgarian_language_pack",
            "remove_bulgarian_language",
        }

    def _is_driver_backup_item(self, item: dict[str, str]) -> bool:
        return item.get("action_id", "") in {
            "driver_backup_clean",
            "driver_backup_full",
            "driver_recovery_usb",
            "driver_pc_report",
            "driver_backup_advanced",
            "driver_restore_last",
        }

    def _is_nexus_admin_item(self, item: dict[str, str]) -> bool:
        return item.get("action_id", "") in {
            "nexus_list_users",
            "nexus_change_password",
            "nexus_create_user",
            "nexus_delete_user",
            "nexus_user_details",
            "nexus_toggle_admin",
        }

    def _office_install_info(self, action_id: str) -> object:
        if action_id not in self.office_inventory_cache:
            self.office_inventory_cache[action_id] = detect_installed_office(action_id)
        return self.office_inventory_cache[action_id]

    def _office_online_status(self, action_id: str) -> object:
        if action_id not in self.office_online_cache:
            self.office_online_cache[action_id] = check_online_package(action_id)
        return self.office_online_cache[action_id]

    def _office_maintenance_status(self, action_id: str) -> object:
        if action_id not in self.office_maintenance_cache:
            self.office_maintenance_cache[action_id] = check_maintenance_action(action_id)
        return self.office_maintenance_cache[action_id]

    def _adobe_reader_status(self) -> object:
        if self.adobe_reader_status_cache is None:
            self.adobe_reader_status_cache = check_adobe_reader_status(PROJECT_ROOT)
        return self.adobe_reader_status_cache

    def _language_status(self) -> object:
        if self.language_status_cache is None:
            self.language_status_cache = get_language_status()
        return self.language_status_cache

    def _reset_language_status_cache(self) -> None:
        self.language_status_cache = None

    def _last_driver_backup_dir(self) -> Path | None:
        last_backup = self.settings.get("last_driver_backup_dir", "")
        if last_backup:
            backup_path = Path(last_backup)
            if backup_path.exists():
                return backup_path
        return None

    def _nexus_admin_status(self) -> object:
        if self.nexus_admin_status_cache is None:
            self.nexus_admin_status_cache = check_nexus_admin_status()
        return self.nexus_admin_status_cache

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
                return "Статусът не може да се провери. Натисни бутона за подробности."

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
                return f"Проверка на езиковите настройки. Налични: {ready_count}/5."
            if action_id == "toggle_bulgarian_bds":
                state = "налична" if language_status.has_bds else "липсва"
                return f"БДС клавиатурна подредба. Статус: {state}."
            if action_id == "toggle_bulgarian_phonetic":
                state = "налична" if language_status.has_phonetic else "липсва"
                return f"Стандартна фонетична подредба. Статус: {state}."
            if action_id == "toggle_bulgarian_traditional":
                state = "налична" if language_status.has_traditional else "липсва"
                return f"Традиционна фонетична подредба. Статус: {state}."
            if action_id == "toggle_bulgarian_language_pack":
                state = "наличен" if language_status.has_language_pack else "липсва"
                return f"Български езиков пакет. Статус: {state}."
            if action_id == "remove_bulgarian_language":
                state = "bg-BG е наличен" if language_status.has_bulgarian else "bg-BG не е намерен"
                return f"Премахване на bg-BG от езиковия списък. Статус: {state}."

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
            extra = ""
            if item.get("action_id") == "nexus_delete_user":
                extra = "\nHigh impact action. A confirmation prompt will appear."
            if item.get("action_id") == "nexus_toggle_admin":
                extra = "\nChanges membership in the local Administrators group."
            return f"{description}\n\n{marker} {nexus_status.message}{extra}"

        return description

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

    def _button_colors(self, kind: str, accent: str) -> tuple[str, str, str]:
        if kind == "menu":
            return ("#1f6fb2", "#f4fbff", "#2b8ddd")
        if kind == "exit":
            return ("#9a2f2f", "#fff6f6", "#c24040")
        if kind == "info":
            return ("#36403a", "#d8e2db", "#36403a")
        return ("#1f8f43", "#f5fff7", "#28b155")

    def _kind_description(self, kind: str) -> str:
        descriptions = {
            "menu": "Open this module and view its available tools.",
            "action": "Prepared action placeholder. We can connect it to a real script next.",
            "exit": "Close the current application session.",
            "info": "Information card for system status or guidance.",
        }
        return descriptions.get(kind, "Module item")

    def _button_text(self, kind: str) -> str:
        labels = {
            "menu": "Enter Menu",
            "action": "Run",
            "exit": "Exit",
            "info": "Info",
        }
        return labels.get(kind, "Open")

    def handle_item(self, item: dict[str, str]) -> None:
        kind = item["kind"]
        if kind == "menu":
            target = item["target"]
            if target == "windows11_activation" and not self._authorize_windows11_menu():
                self.status_var.set("Access to Windows 11 Key Manager was denied.")
                return
            if target != self.current_menu:
                self.history.append(self.current_menu)
            self.render_menu(target)
        elif kind == "action":
            self._handle_action(item)
        elif kind == "exit":
            self.root.destroy()

    def _authorize_windows11_menu(self) -> bool:
        password = simpledialog.askstring(
            "Защитено меню",
            "Въведи парола за менюто за Windows 11:",
            parent=self.root,
            show="*",
        )
        if password is None:
            return False
        expected_hash = self.secure_store.get(
            "admin_menu_password_hash",
            hash_secret(DEFAULT_WINDOWS11_MENU_PASSWORD),
        )
        entered_hash = hash_secret(password)
        default_hash = hash_secret(DEFAULT_WINDOWS11_MENU_PASSWORD)
        if entered_hash == expected_hash:
            return True

        if entered_hash == default_hash:
            self.secure_store["admin_menu_password_hash"] = default_hash
            try:
                save_secure_store(self.secure_store)
            except OSError:
                pass
            return True

        messagebox.showerror("Достъп отказан", "Въведената парола не е правилна.", parent=self.root)
        self.status_var.set("Достъпът до менюто за Windows 11 беше отказан.")
        return False

    def _handle_action(self, item: dict[str, str]) -> None:
        action_id = item.get("action_id", "")
        if action_id == "save_windows11_key":
            self._save_windows11_key()
            return
        if action_id == "activate_windows11":
            self._activate_windows11()
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

        self.status_var.set(f"Selected action: {item['label']}. This is ready to connect to a real Python or PowerShell task.")

    def _center_window(self, window: tk.Toplevel, width: int, height: int) -> None:
        window.update_idletasks()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        position_x = max(0, (screen_width - width) // 2)
        position_y = max(0, (screen_height - height) // 2)
        window.geometry(f"{width}x{height}+{position_x}+{position_y}")

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

    def _save_windows11_key(self) -> None:
        existing_key = self.secure_store.get("windows11_product_key", "")
        prompt = "Enter the Windows 11 product key to save for your admin workflow:"
        product_key = simpledialog.askstring(
            "Save Windows 11 Key",
            prompt,
            parent=self.root,
            initialvalue=existing_key,
        )
        if product_key is None:
            self.status_var.set("Saving Windows 11 product key was canceled.")
            return

        normalized_key = product_key.strip().upper()
        if not normalized_key:
            messagebox.showwarning("Missing Key", "Please enter a product key before saving.", parent=self.root)
            self.status_var.set("No Windows 11 product key was saved.")
            return

        self.secure_store["windows11_product_key"] = normalized_key
        try:
            save_secure_store(self.secure_store)
        except OSError as exc:
            messagebox.showerror(
                "Save Failed",
                f"The Windows 11 product key could not be saved.\n\n{exc}",
                parent=self.root,
            )
            self.status_var.set("Saving the Windows 11 product key failed.")
            return
        self.status_var.set("Windows 11 product key saved successfully.")
        messagebox.showinfo(
            "Key Saved",
            f"The Windows 11 product key has been saved.\n\nSecure file:\n{SECURE_STORE_FILE}",
            parent=self.root,
        )

    def _show_windows11_key(self) -> None:
        saved_key = self.secure_store.get("windows11_product_key", "").strip()
        if not saved_key:
            messagebox.showinfo("No Saved Key", "There is no saved Windows 11 product key yet.", parent=self.root)
            self.status_var.set("No Windows 11 product key is currently stored.")
            return

        messagebox.showinfo("Saved Windows 11 Key", saved_key, parent=self.root)
        self.status_var.set("Displayed the saved Windows 11 product key.")

    def _clear_windows11_key(self) -> None:
        if "windows11_product_key" not in self.secure_store:
            self.status_var.set("There is no saved Windows 11 product key to remove.")
            return

        confirmed = messagebox.askyesno(
            "Clear Saved Key",
            "Do you want to remove the saved Windows 11 product key?",
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set("Saved Windows 11 product key was kept.")
            return

        self.secure_store.pop("windows11_product_key", None)
        try:
            save_secure_store(self.secure_store)
        except OSError as exc:
            messagebox.showerror(
                "Remove Failed",
                f"The saved Windows 11 product key could not be removed.\n\n{exc}",
                parent=self.root,
            )
            self.status_var.set("Removing the saved Windows 11 product key failed.")
            return
        self.status_var.set("Saved Windows 11 product key removed.")

    def _save_office_key(self) -> None:
        selected_action = self._choose_office_version("Save Office Key")
        if not selected_action:
            self.status_var.set("Saving Office product key was canceled.")
            return

        version_label = get_office_version_label(selected_action)
        store_key = f"{selected_action}_product_key"
        existing_key = self.secure_store.get(store_key, "")
        product_key = simpledialog.askstring(
            "Save Office Key",
            f"Enter the product key for {version_label}:",
            parent=self.root,
            initialvalue=existing_key,
        )
        if product_key is None:
            self.status_var.set(f"Saving {version_label} product key was canceled.")
            return

        normalized_key = product_key.strip().upper()
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

    def _activate_windows11(self) -> None:
        saved_key = self.secure_store.get("windows11_product_key", "").strip()
        if not saved_key:
            messagebox.showwarning(
                "Missing Key",
                "Save a Windows 11 product key first, then run activation.",
                parent=self.root,
            )
            self.status_var.set("Windows 11 activation could not start because no key is saved.")
            return

        confirmed = messagebox.askyesno(
            "Activate Windows 11",
            "Run Windows 11 activation now using the saved product key?",
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set("Windows 11 activation was canceled.")
            return

        self.status_var.set("Activating Windows 11...")
        self._open_activation_window(
            title="Windows 11 Activation Progress",
            heading="Windows 11 Activation",
            intro="The application is running the saved activation workflow.",
        )
        threading.Thread(target=self._run_windows11_activation, args=(saved_key,), daemon=True).start()

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

    def _finish_language_action(self, subject: str, success: bool, details: str) -> None:
        self._reset_language_status_cache()
        self._show_activation_result(success, details, subject)
        self._load_language_status_async()
        if self.current_menu == "language":
            self._render_cards()

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

    def _finish_driver_workflow(self, subject: str, success: bool, details: str) -> None:
        self._show_activation_result(success, details, subject)
        if self.current_menu == "driver_backup":
            self._render_cards()

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
            if not username:
                self.status_var.set("User creation was canceled.")
                return
            wants_password = messagebox.askyesno("Create New User", f"Create user {username.strip()} with a password?", parent=self.root)
            password = None
            if wants_password:
                password = simpledialog.askstring("Create New User", f"Enter password for {username.strip()}:", parent=self.root, show="*")
                if password is None:
                    self.status_var.set("User creation was canceled.")
                    return
            make_admin = messagebox.askyesno("Create New User", f"Make {username.strip()} an Administrator?", parent=self.root)
            self._run_nexus_background(
                "Create New User",
                lambda: create_user(username.strip(), password, make_admin),
                subject=f"User {username.strip()}",
            )
            return
        if action_id == "nexus_delete_user":
            username = simpledialog.askstring("Delete User", "Enter the username to delete:", parent=self.root)
            if not username:
                self.status_var.set("User deletion was canceled.")
                return
            confirm_name = simpledialog.askstring(
                "Delete User",
                f'Type the username "{username.strip()}" again to confirm permanent deletion:',
                parent=self.root,
            )
            if confirm_name != username.strip():
                self.status_var.set("User deletion was canceled.")
                return
            self._run_nexus_background("Delete User", lambda: delete_user(username.strip()), subject=f"User {username.strip()}")
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

    def _install_office_offline(self, action_id: str) -> None:
        self._refresh_resource_panel()
        installer = get_office_offline_installer(action_id)
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
            args=(installer,),
            daemon=True,
        ).start()

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

        winget_exe = find_winget_executable()
        if not winget_exe:
            messagebox.showerror(
                "Winget Not Found",
                "Winget is required for online Office installation but was not found.",
                parent=self.root,
            )
            self.status_var.set("Online Office installation could not start because winget is missing.")
            return

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
            args=(package.label, package.winget_id, winget_exe),
            daemon=True,
        ).start()

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

        confirmed = messagebox.askyesno(
            "Adobe Reader",
            (
                f"{details}\n\n"
                "Да инсталирам/обновя ли Adobe Reader до актуалната версия чрез winget?"
            ),
            parent=self.root,
        )
        if not confirmed:
            self.status_var.set("Adobe Reader инсталацията беше отказана.")
            return

        self.status_var.set("Стартиране на Adobe Reader online инсталация...")
        self._open_activation_window(
            title="Adobe Reader",
            heading="Adobe Reader Online Setup",
            intro="Проверява се актуалната версия и се стартира инсталация чрез winget.",
        )
        threading.Thread(
            target=self._run_adobe_reader_installation,
            args=(winget_exe,),
            daemon=True,
        ).start()

    def _run_adobe_reader_installation(self, winget_exe: str) -> None:
        command = [
            winget_exe,
            "install",
            "--id",
            ADOBE_READER_WINGET_ID,
            "--source",
            "winget",
            "--silent",
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
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    55,
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

    def _run_office_offline_installation(self, installer: object) -> None:
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
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    45,
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

    def _run_office_online_installation(self, label: str, winget_id: str, winget_exe: str) -> None:
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
            self.root.after(
                0,
                lambda: self._update_activation_progress(
                    45,
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

    def _finish_office_installation(self, action_id: str, subject: str, success: bool, details: str) -> None:
        self.office_inventory_cache.pop(action_id, None)
        self._show_activation_result(success, details, subject)
        if self.current_menu == "office_install_center":
            self._render_cards()

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

    def _finish_office_removal(self, action_id: str, subject: str, success: bool, details: str) -> None:
        self.office_inventory_cache.pop(action_id, None)
        self._show_activation_result(success, details, subject)
        if self.current_menu == "office_install_center":
            self._render_cards()

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

    def _run_windows11_activation(self, product_key: str) -> None:
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
            self.root.after(0, lambda: self._update_activation_progress(10, "Preparing activation environment...", "Starting Windows 11 activation workflow."))
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
            self.root.after(0, lambda: self._show_activation_result(False, str(exc), "Windows 11"))
            return

        final_output = "\n\n".join(output_lines) or "Windows 11 activation completed successfully."
        self.root.after(0, lambda: self._show_activation_result(True, final_output, "Windows 11"))

    def _show_activation_result(self, success: bool, details: str, subject: str) -> None:
        if success:
            self.status_var.set(f"{subject} activation completed.")
            self._update_activation_progress(100, "Activation completed.", details, finished=True, success=True)
            return

        self.status_var.set(f"{subject} activation failed.")
        self._update_activation_progress(100, "Activation failed.", details, finished=True, success=False)

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

    def _append_activation_log(self, text: str) -> None:
        if self.activation_log_widget is None or not self.activation_log_widget.winfo_exists():
            return
        self.activation_log_widget.config(state="normal")
        self.activation_log_widget.insert("end", f"{text}\n\n")
        self.activation_log_widget.see("end")
        self.activation_log_widget.config(state="disabled")

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

    def go_back(self) -> None:
        if not self.history:
            return
        target = self.history.pop()
        self.current_menu = target
        self.current_page = 0
        self.menu_path.config(text=self._build_path())
        self.card_title.config(text=MENU_TREE[target]["title"])
        self.card_subtitle.config(text=MENU_TREE[target]["subtitle"])
        self.subtitle_label.config(text=MENU_TREE[target]["subtitle"])
        self.status_var.set(f"Returned to {MENU_TREE[target]['title']}.")
        self.header_home_button.config(state="disabled" if target == "main" else "normal")
        self._render_cards()

    def go_home(self) -> None:
        self.render_menu("main", reset_history=True)

    def next_page(self) -> None:
        items = MENU_TREE[self.current_menu]["items"]
        page_size = MENU_PAGE_SIZE.get(self.current_menu, CARDS_PER_PAGE)
        total_pages = max(1, math.ceil(len(items) / page_size))
        if self.current_page < total_pages - 1:
            self.current_page += 1
            self._render_cards()

    def previous_page(self) -> None:
        if self.current_page > 0:
            self.current_page -= 1
            self._render_cards()


def main() -> None:
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
    SplashScreen(root)
    root.mainloop()


if __name__ == "__main__":
    main()
