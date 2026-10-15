from __future__ import annotations

import subprocess
import sys

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QDialog, QHBoxLayout, QLabel, QMainWindow, QMessageBox, QPushButton, QStackedWidget, QStatusBar, QVBoxLayout, QWidget

from nexus_admin import change_password, check_nexus_admin_status, create_user, delete_user, list_users, set_admin_rights, user_details
from pages.execution_page import ExecutionProgressPage
from pages.home import HomePage, WgaDashboardPage
from pages.simple_page import LegacyMenuPage, SectionAction, SimpleSectionPage
from pages.software_page import SoftwarePage
from services.admin import relaunch_as_admin
from services.action_runner import format_action_history
from services.key_store import (
    WINDOWS_KEYS,
    clear_key,
    get_key,
    normalize_product_key,
    office_key_choices,
    office_key_label,
    office_store_key,
    save_key,
    secure_store_path,
    windows_key_choices,
    windows_key_label,
    windows_store_key,
)
from services.paths import project_root, qss_path
from ui.dialogs import modern_choice, modern_confirm, modern_details_message, modern_message, modern_text_input
from ui.loading_screen import LoadingScreen
from ui.right_panel import RightPanel
from ui.sidebar import Sidebar


class MainWindow(QMainWindow):
    PAGE_META = {
        "windows": {
            "title": "Windows",
            "description": "Активация, ключове, desktop настройки и системни Windows действия.",
            "legacy": "activation",
            "actions": [
                SectionAction("Windows 10 активация", "Запис на ключ, показване/изчистване на ключ и стартиране на активация за Windows 10.", "windows10_activation", "features/windows_10_11.png", "#168CFF"),
                SectionAction("Windows 11 активация", "Същият activation workflow, отделен за Windows 11 и новите build версии.", "windows11_activation", "quick/windows.png", "#20E28A"),
            ],
        },
        "office": {
            "title": "Office",
            "description": "Office активация, offline/online инсталация, status проверка, repair и cleanup.",
            "legacy": "office_activation",
            "actions": [
                SectionAction("Office активация", "Управление на Office ключ и активация за Office 2016, 2019 и 2021.", "office_activation", "features/office_2016_2024.png", "#FF4B1F"),
                SectionAction("Office Install Center", "Обединява offline инсталациите за Office 2016, 2019, 2021 и 2024.", "office_install_center", "quick/office.png", "#7657FF"),
                SectionAction("Office Online God Mode", "Online deployment пакети, status проверка, repair и принудително премахване на Office.", "office_center", "tools/onedrive.png", "#20C5FF"),
            ],
        },
        "drivers": {
            "title": "Драйвери",
            "description": "Backup, restore, PC report и recovery USB инструменти.",
            "legacy": "driver_backup",
            "actions": [
                SectionAction("Driver Backup", "Чист или пълен backup на драйверите през наличния WGA workflow.", "driver_backup", "tools/drivers.png", "#7657FF"),
                SectionAction("Recovery USB", "Създаване на recovery USB с restore batch за драйвери.", "driver_backup", "tools/recovery_usb.png", "#20E28A"),
                SectionAction("PC отчет", "Генериране на хардуерен отчет с CPU, RAM, GPU, BIOS, дискове и мрежа.", "driver_backup", "actions/pc_report.png", "#20C5FF"),
            ],
        },
        "network": {
            "title": "Мрежа и интернет",
            "description": "Мрежови настройки, диагностика и Internet related инструменти.",
            "legacy": "main",
            "actions": [
                SectionAction("Мрежови инструменти", "Отваря WGA hub с наличните network и system действия.", "main", "tools/network.png", "#20C5FF"),
                SectionAction("PC отчет с мрежа", "Хардуерен и network отчет през driver/report модула.", "driver_backup", "actions/pc_report.png", "#168CFF"),
                SectionAction("Разширени проверки", "Hidden status и agent информация за диагностика.", "hidden_menu", "navigation/advanced.png", "#7657FF"),
            ],
        },
        "system": {
            "title": "Система",
            "description": "Системна информация, health проверки, desktop настройки и административни инструменти.",
            "legacy": "main",
            "actions": [
                SectionAction("Системен hub", "Централно меню за системни действия, desktop icons и admin tools.", "main", "tools/system_tools.png", "#168CFF"),
                SectionAction("Nexus Admin", "Локални потребители, администраторски акаунт и admin extras.", "nexus_admin", "navigation/system.png", "#20E28A"),
                SectionAction("Hidden status", "Agent status и скрита диагностична информация.", "hidden_menu", "navigation/advanced.png", "#7657FF"),
            ],
        },
        "tools": {
            "title": "Инструменти",
            "description": "Инсталации, Office deployment, Ninite, Adobe Reader, OneDrive и WGA помощни действия.",
            "legacy": "install_software",
            "actions": [
                SectionAction("Install Software", "Office Install Center, Ninite, Adobe Reader и secret install interface.", "install_software", "tools/services.png", "#168CFF"),
                SectionAction("Auto Installer", "Автоматизиран installer hub за пакетни инсталации.", "auto_installer", "tools/startup.png", "#20E28A"),
                SectionAction("Инсталационни ресурси", "Проверка на Installers manifest, липсващи пакети и готовност на локалните ресурси.", "execute::resources_status", "actions/open.png", "#20C5FF"),
                SectionAction("Изтегли липсващи ресурси", "Изтегля липсващите пакети от installers_manifest.json и ги разархивира при нужда.", "execute::resources_download_missing", "tools/startup.png", "#FFB000"),
                SectionAction("Office Deployment", "Online Office deployment, repair, status и cleanup инструменти.", "office_center", "quick/office.png", "#FF4B1F"),
                SectionAction("OneDrive", "Reset и управление на OneDrive workflows.", "reset_onedrive", "tools/onedrive.png", "#20C5FF"),
                SectionAction("GOD MODE", "Разширено меню със скрити и advanced WGA действия.", "hidden_menu", "actions/god_mode.png", "#7657FF"),
            ],
        },
        "personalization": {
            "title": "Персонализация",
            "description": "Език, клавиатури, desktop и Windows персонализация.",
            "legacy": "language",
            "actions": [
                SectionAction("Езиков мениджър", "Bulgarian BDS, phonetic, traditional, language pack и refresh.", "language", "tools/personalization.png", "#7657FF"),
                SectionAction("Desktop icons", "Показване на системните икони на работния плот.", "main", "navigation/personalization.png", "#168CFF"),
                SectionAction("Windows настройки", "Windows системни настройки и помощни действия.", "main", "quick/windows.png", "#20C5FF"),
            ],
        },
        "privacy": {
            "title": "Защита и поверителност",
            "description": "Security действия, privacy настройки и диагностични проверки.",
            "legacy": "main",
            "actions": [
                SectionAction("Privacy hub", "Основен WGA hub със security/privacy свързани действия.", "main", "tools/privacy.png", "#20E28A"),
                SectionAction("Hidden status", "Проверки на agent status и системни диагностични данни.", "hidden_menu", "status/ready.png", "#20C5FF"),
                SectionAction("Nexus Admin", "Администраторски акаунти и локални user настройки.", "nexus_admin", "navigation/system.png", "#7657FF"),
            ],
        },
        "advanced": {
            "title": "Разширени",
            "description": "Hidden menu, Nexus Admin, advanced install и диагностични WGA инструменти.",
            "legacy": "hidden_menu",
            "actions": [
                SectionAction("Hidden Agent Menu", "Показване и зареждане на hidden agent status.", "hidden_menu", "navigation/advanced.png", "#7657FF"),
                SectionAction("Nexus Admin", "Локални потребители и admin инструменти.", "nexus_admin", "tools/system_tools.png", "#168CFF"),
                SectionAction("Secret Install", "Разширени install presets за runtimes, utilities и engineering stacks.", "secret_install", "actions/god_mode.png", "#FFB000"),
            ],
        },
        "settings": {
            "title": "Настройки",
            "description": "Настройки на WGA, версия, update поведение, логове и история на операциите.",
            "legacy": "main",
            "actions": [
                SectionAction("История на операциите", "Показва последните стартирани WGA действия, статус, прогрес и подробности от лога.", "operation_history", "navigation/information.png", "#20E28A"),
                SectionAction("Основни настройки", "Отваря основния WGA раздел за наличните настройки.", "main", "navigation/settings.png", "#168CFF"),
                SectionAction("Информация", "Версия, компоненти и информация за приложението.", "main", "navigation/information.png", "#20C5FF"),
                SectionAction("Advanced", "Скрити и диагностични настройки.", "hidden_menu", "navigation/advanced.png", "#7657FF"),
            ],
        },
        "about": {
            "title": "Информация",
            "description": "Версия, компоненти, системен статус и диагностични WGA данни.",
            "legacy": "main",
            "actions": [
                SectionAction("WGA main hub", "Главният WGA раздел с наличните модули.", "main", "logo/wga_symbol.png", "#168CFF"),
                SectionAction("PC отчет", "Генериране на системен отчет за машината.", "driver_backup", "actions/pc_report.png", "#20E28A"),
                SectionAction("Hidden status", "Диагностичен status на agent и скрити модули.", "hidden_menu", "status/ready.png", "#7657FF"),
            ],
        },
    }

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("WinSys Guardian Advanced")
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint, True)
        self.resize(1240, 760)
        self.setMinimumSize(980, 620)
        self._load_styles()
        central = QWidget()
        central.setObjectName("centralRoot")
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        self.sidebar = Sidebar()
        self.sidebar.section_selected.connect(self.show_page)
        self.sidebar.exit_requested.connect(self.close)
        root.addWidget(self.sidebar)
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(14, 14, 14, 10)
        content_layout.setSpacing(10)
        header_row = QHBoxLayout()
        self.header = QLabel("Начало")
        self.header.setProperty("title", True)
        header_row.addWidget(self.header)
        header_row.addStretch(1)
        close_button = QPushButton("X")
        close_button.setProperty("windowClose", True)
        close_button.setFixedSize(42, 34)
        close_button.clicked.connect(self.close)
        header_row.addWidget(close_button)
        content_layout.addLayout(header_row)
        body = QHBoxLayout()
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(14)
        self.stack = QStackedWidget()
        self.pages: dict[str, int] = {}
        self.execution_dialogs: list[QDialog] = []
        self.active_execution_actions: dict[str, QDialog] = {}
        self.module_processes: dict[str, subprocess.Popen] = {}
        self.loading_screens: list[LoadingScreen] = []
        self._opening_wga = False
        home = HomePage()
        home.navigate.connect(self.show_page)
        self._add_page("home", home)
        wga_dashboard = WgaDashboardPage()
        wga_dashboard.navigate.connect(self.show_page)
        self._add_page("wga_dashboard", wga_dashboard)
        for key, meta in self.PAGE_META.items():
            page = SimpleSectionPage(meta["title"], meta["description"], meta["legacy"], meta["actions"])
            page.navigate.connect(self.show_page)
            self._add_page(key, page)
        self._add_page(
            "execute::auto_installer",
            ExecutionProgressPage("Инсталиране на приложения", "Автоматичен инсталатор с детайлен progress.", "auto_installer", mode="installer"),
        )
        self._add_page(
            "execute::resources_status",
            ExecutionProgressPage("Инсталационни ресурси", "Проверява manifest-а, наличните файлове и липсващите installer пакети.", "resources_status"),
        )
        self._add_page(
            "execute::resources_download_missing",
            ExecutionProgressPage("Изтегляне на ресурси", "Изтегля липсващите installer пакети от manifest-а и ги разархивира при нужда.", "resources_download_missing"),
        )
        self._add_legacy_menu_pages()
        software_page = SoftwarePage()
        self._add_page("software", software_page)
        body.addWidget(self.stack, 1)
        self.right_panel = RightPanel()
        self.right_panel.navigate.connect(self.show_page)
        body.addWidget(self.right_panel)
        content_layout.addLayout(body, 1)
        root.addWidget(content, 1)
        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("Готов за работа | Windows: проверява се | Office: проверява се | Драйвери: налични инструменти | OneDrive: налични инструменти")
        self.show_page("home")

    def _add_page(self, key: str, widget: QWidget) -> None:
        self.pages[key] = self.stack.addWidget(widget)

    def _add_legacy_menu_pages(self) -> None:
        try:
            from legacy_app import MENU_TREE
        except Exception:
            MENU_TREE = {}
        parent_by_key: dict[str, str] = {}
        for parent_key, parent_data in MENU_TREE.items():
            if not isinstance(parent_data, dict):
                continue
            for item in parent_data.get("items", []):
                if isinstance(item, dict) and item.get("kind") == "menu" and item.get("target"):
                    parent_by_key.setdefault(str(item["target"]), str(parent_key))
        for key, data in MENU_TREE.items():
            if key in self.pages or not isinstance(data, dict):
                continue
            if key == "auto_installer":
                page = SimpleSectionPage(
                    "Автоматичен инсталатор",
                    "Инсталиране на избрани приложения с детайлно проследяване по стъпки.",
                    "execute::auto_installer",
                    [
                        SectionAction("Инсталиране на приложения", "Отваря новия progress екран с детайлен статус за всяко приложение.", "execute::auto_installer", "tools/startup.png", "#20E28A"),
                        SectionAction("Office Deployment", "Подготовка за Office deployment задачи в новия интерфейс.", "office_center", "quick/office.png", "#7657FF"),
                        SectionAction("Install Software", "Допълнителни install менюта и приложения.", "install_software", "tools/services.png", "#168CFF"),
                    ],
                )
                page.navigate.connect(self.show_page)
                self._add_page(key, page)
                continue
            data = self._modern_menu_data(key, data)
            page = LegacyMenuPage(key, data, parent_by_key.get(key))
            page.navigate.connect(self.show_page)
            self._add_page(key, page)
            for item in data.get("items", []):
                if not isinstance(item, dict) or item.get("kind") != "action":
                    continue
                execute_key = self._execution_key(key, item)
                if execute_key in self.pages:
                    continue
                self._add_page(
                    execute_key,
                    ExecutionProgressPage(
                        str(item.get("label") or "WGA действие"),
                        str(item.get("description") or "Изпълнение на WGA действие."),
                        str(item.get("action_id") or item.get("label") or "wga_action"),
                    ),
                )

    def _modern_menu_data(self, key: str, data: dict[str, object]) -> dict[str, object]:
        if key == "main":
            return self._main_modern_menu_data(data)
        if key == "secret_install":
            return self._secret_install_modern_menu_data(data)
        if key == "office_activation":
            return self._office_modern_menu_data(data)
        if key not in {"windows10_activation", "windows11_activation"}:
            return data
        items = list(data.get("items", []))
        if not any(isinstance(item, dict) and item.get("action_id") == "auto_windows_activation" for item in items):
            items.insert(
                0,
                {
                    "label": "Автоматично активиране",
                    "kind": "action",
                    "action_id": "auto_windows_activation",
                    "description": "Проверява Windows версията, build номера и записания ключ, после избира правилния activation workflow.",
                },
            )
        modern_data = dict(data)
        modern_data["items"] = items
        return modern_data

    def _main_modern_menu_data(self, data: dict[str, object]) -> dict[str, object]:
        items = []
        for item in data.get("items", []):
            if isinstance(item, dict) and item.get("kind") == "action" and not item.get("action_id"):
                item = dict(item)
                item["action_id"] = "reset_console"
                item["label"] = "Опресни интерфейса"
            items.append(item)
        modern_data = dict(data)
        modern_data["items"] = items
        return modern_data

    def _secret_install_modern_menu_data(self, data: dict[str, object]) -> dict[str, object]:
        preset_ids = [
            "system_runtimes",
            "browsers_comms",
            "development",
            "languages_db",
            "cybersecurity_net",
            "virtualization",
            "multimedia_design",
            "gaming_tools",
            "utilities_office",
            "advanced_tools",
            "update_all_apps",
        ]
        items = []
        action_index = 0
        for item in data.get("items", []):
            if isinstance(item, dict) and item.get("kind") == "action" and not item.get("action_id"):
                item = dict(item)
                if action_index < len(preset_ids):
                    item["action_id"] = preset_ids[action_index]
                action_index += 1
            items.append(item)
        modern_data = dict(data)
        modern_data["items"] = items
        return modern_data

    def _office_modern_menu_data(self, data: dict[str, object]) -> dict[str, object]:
        items = list(data.get("items", []))
        if not any(isinstance(item, dict) and item.get("action_id") == "auto_office_activation" for item in items):
            items.insert(
                0,
                {
                    "label": "Автоматично активиране",
                    "kind": "action",
                    "action_id": "auto_office_activation",
                    "description": "Открива инсталирания Office пакет, проверява дали има записан ключ за него и стартира правилния activation workflow.",
                },
            )
        modern_data = dict(data)
        modern_data["items"] = items
        return modern_data

    def show_page(self, key: str) -> None:
        if key == "install_software":
            key = "software"
        if key == "app_exit":
            self.close()
            return
        if key == "app_minimize":
            self.showMinimized()
            return
        if key == "app_toggle_maximize":
            if self.isMaximized():
                self.showNormal()
            else:
                self.showMaximized()
            return
        if key == "wga_dashboard" and self.stack.currentIndex() == self.pages.get("home") and not self._opening_wga:
            self._open_wga_with_loading()
            return
        if key.startswith("module::"):
            self._open_external_module(key.removeprefix("module::"))
            return
        if key == "operation_history":
            self._show_operation_history()
            return
        if key.startswith("execute::"):
            self._open_execution_dialog(key)
            return
        if key not in self.pages:
            key = "home"
        self.stack.setCurrentIndex(self.pages[key])
        chrome_visible = key != "home"
        self.sidebar.setVisible(chrome_visible)
        self.right_panel.setVisible(chrome_visible)
        self.header.setText("WinSys Guardian Advanced" if key == "wga_dashboard" else self.PAGE_META.get(key, {}).get("title", self._page_title(key)))
        self.header.setVisible(chrome_visible)
        self.statusBar().setVisible(chrome_visible)
        self.sidebar.set_selected("home" if key == "wga_dashboard" else key)
        if chrome_visible:
            self._show_wga_fullscreen()

    def _show_wga_fullscreen(self) -> None:
        self.setWindowState(self.windowState() | Qt.WindowState.WindowMaximized)
        if self.isVisible():
            self.showMaximized()

    def _open_wga_with_loading(self) -> None:
        self._opening_wga = True
        loading = LoadingScreen()
        self.loading_screens.append(loading)
        self.hide()

        def handle_failure(detail: str) -> None:
            self._opening_wga = False
            self.loading_screens.remove(loading) if loading in self.loading_screens else None
            if "администраторски права" in detail.lower():
                if relaunch_as_admin():
                    self.close()
                    return
                QMessageBox.critical(self, "Administrator Rights Required", "WinSys Guardian Advanced must run with administrator rights.")
            else:
                QMessageBox.critical(self, "WGA Startup Error", f"WGA не можа да зареди:\n\n{detail}")
            self.show()
            self.raise_()
            self.activateWindow()

        def handle_completed(warnings: object) -> None:
            loading.complete_gui_step("interface", "Създаване на WGA интерфейса.", 92)
            self.show_page("wga_dashboard")
            self._opening_wga = False
            loading.complete_gui_step("launcher", "WGA началният панел е готов.", 98)
            if warnings:
                self.statusBar().showMessage("WGA е зареден с предупреждения. Провери logs/wga-startup.log.")
            loading.finish_success(self)

        def cleanup() -> None:
            if loading in self.loading_screens:
                self.loading_screens.remove(loading)

        loading.completed.connect(handle_completed)
        loading.failed.connect(handle_failure)
        loading.destroyed.connect(cleanup)
        loading.resize(1100, 680)
        frame = loading.frameGeometry()
        screen = self.screen() or QApplication.primaryScreen()
        if screen is not None:
            frame.moveCenter(screen.availableGeometry().center())
        else:
            frame.moveCenter(self.frameGeometry().center())
        loading.move(frame.topLeft())
        loading.show()
        loading.raise_()
        loading.activateWindow()
        loading.start()

    def _open_external_module(self, module_id: str) -> None:
        if module_id not in {"network", "optimization"}:
            modern_message(self, "Модулът не е намерен", f"Няма регистриран старт за '{module_id}'.")
            return
        existing = self.module_processes.get(module_id)
        if existing is not None and existing.poll() is None:
            modern_message(self, "Модулът вече е отворен", "Вече има активен прозорец за този модул.")
            return
        try:
            if getattr(sys, "frozen", False):
                args = [sys.executable, f"--{module_id}"]
            else:
                args = [sys.executable, str(project_root() / "app.py"), f"--{module_id}"]
            self.module_processes[module_id] = subprocess.Popen(
                args,
                cwd=str(project_root()),
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            self.statusBar().showMessage(f"Отворен модул: {module_id}")
        except Exception as exc:
            modern_message(self, "Грешка при отваряне", f"Модулът не можа да се стартира:\n{exc}")

    def _page_title(self, key: str) -> str:
        if key == "software":
            return "Софтуер и пакети"
        widget = self.stack.widget(self.pages[key])
        if isinstance(widget, LegacyMenuPage):
            return widget.menu_key.replace("_", " ").title()
        if isinstance(widget, ExecutionProgressPage):
            return widget.title_text
        return key

    def _open_execution_dialog(self, key: str) -> None:
        if key not in self.pages:
            return
        template = self.stack.widget(self.pages[key])
        if not isinstance(template, ExecutionProgressPage):
            return
        active_dialog = self.active_execution_actions.get(template.action_id)
        if active_dialog is not None and active_dialog.isVisible():
            active_dialog.raise_()
            active_dialog.activateWindow()
            modern_message(self, "Операцията вече е отворена", f"'{template.title_text}' вече има активен прозорец.")
            return
        if self._handle_activation_state_action(template.action_id):
            return
        if self._handle_nexus_action(template.action_id):
            return
        if template.action_id == "office_force_uninstall_all":
            confirm_text, ok = modern_text_input(
                self,
                "Force Uninstall Office",
                "Това ще опита да премахне всички Office пакети на този компютър.\n\nВъведи CONFIRM, за да продължиш:",
            )
            if not ok or confirm_text.strip().upper() != "CONFIRM":
                self.statusBar().showMessage("Force uninstall Office беше отказан.")
                return
        if template.action_id == "resources_download_missing" and not modern_confirm(
            self,
            "Изтегляне на ресурси",
            "Това може да изтегли големи installer пакети според installers_manifest.json. Да продължа ли?",
        ):
            self.statusBar().showMessage("Изтеглянето на ресурси беше отказано.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle(template.title_text)
        dialog.setWindowFlag(Qt.WindowType.FramelessWindowHint, True)
        dialog.setModal(False)
        dialog.resize(980, 760)
        dialog.setMinimumSize(860, 660)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(10, 10, 10, 10)
        chrome = QHBoxLayout()
        title = QLabel(template.title_text)
        title.setProperty("subtitle", True)
        chrome.addWidget(title)
        chrome.addStretch(1)
        page = ExecutionProgressPage(template.title_text, template.description_text, template.action_id, template.mode)
        close_button = QPushButton("X")
        close_button.setProperty("secondary", True)
        close_button.setFixedWidth(42)
        close_button.clicked.connect(lambda _checked=False, current_page=page, current_dialog=dialog: self._close_execution_dialog(current_page, current_dialog))
        chrome.addWidget(close_button)
        layout.addLayout(chrome)
        page.close_requested.connect(dialog.close)
        layout.addWidget(page)
        self.active_execution_actions[template.action_id] = dialog
        dialog.finished.connect(lambda _result, current=dialog, action_id=template.action_id: self._forget_execution_dialog(current, action_id))
        self.execution_dialogs.append(dialog)
        dialog.show()
        if template.mode != "installer":
            page.start()

    def _forget_execution_dialog(self, dialog: QDialog, action_id: str | None = None) -> None:
        if dialog in self.execution_dialogs:
            self.execution_dialogs.remove(dialog)
        if action_id and self.active_execution_actions.get(action_id) is dialog:
            self.active_execution_actions.pop(action_id, None)

    def _close_execution_dialog(self, page: ExecutionProgressPage, dialog: QDialog) -> None:
        if page.thread and page.thread.isRunning():
            modern_message(self, "Операцията още работи", "Изчакай действието да приключи. Прозорецът може да се затвори, когато бутонът стане 'Затвори'.")
            return
        dialog.close()

    def _handle_activation_state_action(self, action_id: str) -> bool:
        if action_id in WINDOWS_KEYS:
            version_label, fallback_store_key = WINDOWS_KEYS[action_id]
            selected_edition = self._choose_windows_edition(version_label, action_id)
            if not selected_edition:
                return True
            version_label = windows_key_label(version_label, selected_edition)
            store_key = windows_store_key(WINDOWS_KEYS[action_id][0], selected_edition)
            if selected_edition == "Общ ключ":
                store_key = fallback_store_key
            if action_id.startswith("save_"):
                self._save_product_key(version_label, store_key)
            elif action_id.startswith("show_"):
                self._show_product_key(version_label, store_key)
            else:
                self._clear_product_key(version_label, store_key)
            return True
        if action_id in {"save_office_key", "show_office_key", "clear_office_key"}:
            selected = self._choose_office_product(action_id)
            if not selected:
                return True
            version_label = office_key_label(selected)
            store_key = office_store_key(selected)
            if action_id == "save_office_key":
                self._save_product_key(version_label, store_key)
            elif action_id == "show_office_key":
                self._show_product_key(version_label, store_key)
            else:
                self._clear_product_key(version_label, store_key)
            return True
        return False

    def _handle_nexus_action(self, action_id: str) -> bool:
        if not action_id.startswith("nexus_"):
            return False
        status = check_nexus_admin_status()
        if not status.available:
            modern_message(self, "Nexus Admin", status.message)
            return True
        try:
            if action_id == "nexus_list_users":
                result = list_users()
                self._show_command_result("Локални потребители", [result])
            elif action_id == "nexus_user_details":
                username = self._ask_username("Детайли за потребител")
                if username:
                    self._show_command_result(f"Детайли: {username}", [user_details(username)])
            elif action_id == "nexus_change_password":
                username = self._ask_username("Смяна на парола")
                if not username:
                    return True
                password, ok = modern_text_input(self, "Нова парола", f"Въведи нова парола за {username}:", password=True)
                if ok and password:
                    self._show_command_result(f"Парола: {username}", [change_password(username, password)], hide_sensitive=password)
            elif action_id == "nexus_create_user":
                username = self._ask_username("Нов потребител")
                if not username:
                    return True
                password, ok = modern_text_input(self, "Парола", f"Въведи парола за {username} или остави празно:", password=True)
                if not ok:
                    return True
                make_admin = modern_confirm(self, "Admin права", f"Да добавя ли {username} в Administrators?")
                self._show_command_result(f"Създаване: {username}", create_user(username, password, make_admin), hide_sensitive=password)
            elif action_id == "nexus_delete_user":
                username = self._ask_username("Изтриване на потребител")
                if username and modern_confirm(self, "Изтриване", f"Да изтрия ли локалния потребител {username}?"):
                    self._show_command_result(f"Изтриване: {username}", [delete_user(username)])
            elif action_id == "nexus_toggle_admin":
                username = self._ask_username("Administrator права")
                if not username:
                    return True
                choice, ok = modern_choice(self, "Administrator права", f"Избери действие за {username}:", ["Добави admin права", "Премахни admin права"])
                if ok:
                    self._show_command_result(choice, [set_admin_rights(username, choice.startswith("Добави"))])
        except Exception as exc:
            modern_message(self, "Nexus Admin грешка", str(exc))
        return True

    def _ask_username(self, title: str) -> str:
        username, ok = modern_text_input(self, title, "Въведи локално потребителско име:")
        return username.strip() if ok else ""

    def _show_command_result(self, title: str, results, hide_sensitive: str = "") -> None:
        lines: list[str] = []
        success = True
        for result in results:
            success = success and result.returncode == 0
            output = "\n".join(part.strip() for part in (result.stdout, result.stderr) if part and part.strip())
            if hide_sensitive:
                output = output.replace(hide_sensitive, "[hidden]")
            lines.append(output or f"Командата завърши с код {result.returncode}.")
        modern_message(self, title if success else f"{title} - грешка", "\n\n".join(lines))
        self.statusBar().showMessage(f"{title}: {'готово' if success else 'има грешка'}.")

    def _show_operation_history(self) -> None:
        details = format_action_history(120)
        modern_details_message(
            self,
            "История на операциите",
            "Последни WGA действия от локалния лог. Тук се виждат модулът, статусът, прогресът и детайлите.",
            details,
        )
        self.statusBar().showMessage("Показана е историята на операциите.")

    def _choose_windows_edition(self, version_label: str, action_id: str) -> str | None:
        title = {
            "save_windows10_key": "Запази Windows 10 ключ",
            "show_windows10_key": "Покажи Windows 10 ключ",
            "clear_windows10_key": "Изчисти Windows 10 ключ",
            "save_windows11_key": "Запази Windows 11 ключ",
            "show_windows11_key": "Покажи Windows 11 ключ",
            "clear_windows11_key": "Изчисти Windows 11 ключ",
        }.get(action_id, f"{version_label} ключ")
        selected, ok = modern_choice(self, title, "Избери edition/тип ключ:", windows_key_choices())
        if not ok:
            return None
        return selected

    def _choose_office_product(self, action_id: str) -> str | None:
        title = {
            "save_office_key": "Запази Office ключ",
            "show_office_key": "Покажи Office ключ",
            "clear_office_key": "Изчисти Office ключ",
        }.get(action_id, "Office ключ")
        selected, ok = modern_choice(self, title, "Избери Office продукт/версия:", office_key_choices())
        if not ok:
            return None
        return selected

    def _save_product_key(self, version_label: str, store_key: str) -> None:
        existing_key = get_key(store_key)
        raw_key, ok = modern_text_input(
            self,
            f"Запази {version_label} ключ",
            f"Въведи продуктов ключ за {version_label}:",
            existing_key,
        )
        if not ok:
            self.statusBar().showMessage(f"Записът на {version_label} ключ беше отказан.")
            return
        normalized = normalize_product_key(raw_key)
        if not normalized:
            modern_message(self, "Невалиден ключ", "Ключът трябва да съдържа 25 символа. Може да е с тирета или без.")
            self.statusBar().showMessage(f"Не е записан ключ за {version_label}.")
            return
        try:
            save_key(store_key, normalized)
        except OSError as exc:
            modern_message(self, "Грешка при запис", f"Ключът не беше записан.\n\n{exc}")
            return
        modern_message(self, "Ключът е записан", f"{version_label} ключът е записан успешно.\n\nSecure store:\n{secure_store_path()}")
        self.statusBar().showMessage(f"{version_label} ключът е записан.")

    def _show_product_key(self, version_label: str, store_key: str) -> None:
        saved_key = get_key(store_key)
        if not saved_key:
            modern_message(self, "Няма записан ключ", f"Няма записан ключ за {version_label}.")
            self.statusBar().showMessage(f"Няма записан ключ за {version_label}.")
            return
        modern_message(self, f"Записан {version_label} ключ", saved_key)
        self.statusBar().showMessage(f"Показан е записаният ключ за {version_label}.")

    def _clear_product_key(self, version_label: str, store_key: str) -> None:
        if not get_key(store_key):
            modern_message(self, "Няма записан ключ", f"Няма записан ключ за {version_label}.")
            self.statusBar().showMessage(f"Няма записан ключ за {version_label}.")
            return
        if not modern_confirm(self, "Изтриване на ключ", f"Да изтрия ли записания ключ за {version_label}?"):
            self.statusBar().showMessage(f"Ключът за {version_label} е запазен.")
            return
        try:
            clear_key(store_key)
        except OSError as exc:
            modern_message(self, "Грешка при изтриване", f"Ключът не беше изтрит.\n\n{exc}")
            return
        modern_message(self, "Ключът е изтрит", f"Записаният ключ за {version_label} беше изтрит.")
        self.statusBar().showMessage(f"Ключът за {version_label} беше изтрит.")

    def _execution_key(self, menu_key: str, item: dict[str, object]) -> str:
        import re

        label = re.sub(r"[^a-z0-9]+", "_", str(item.get("label") or "action").strip().lower()).strip("_") or "action"
        action_id = re.sub(r"[^a-z0-9]+", "_", str(item.get("action_id") or label).strip().lower()).strip("_") or "action"
        return f"execute::{menu_key}::{action_id}"

    def _load_styles(self) -> None:
        path = qss_path()
        if path.exists():
            self.setStyleSheet(path.read_text(encoding="utf-8"))
