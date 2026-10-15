from __future__ import annotations

import json
import sys
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QObject, QRectF, Qt, QThread, QTimer, Signal, Slot
from PySide6.QtGui import QColor, QPainter, QPainterPath, QPixmap
from PySide6.QtWidgets import (
    QFrame,
    QGraphicsDropShadowEffect,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from services.paths import ui_asset
from services.paths import project_root
from services.paths import start_menu_asset
from services.paths import runtime_file
from services.version import current_version
from self_updater import launch_helper_and_exit, prepare_update_install
from ui.dialogs import modern_message, modern_update_prompt
from update_checker import UpdateResult, check_for_updates
from widgets.cards import FeatureCard, QuickActionCard, ToolCard


class UpdateCheckWorker(QObject):
    finished = Signal(object)

    @Slot()
    def run(self) -> None:
        try:
            data = json.loads(runtime_file("version.json").read_text(encoding="utf-8-sig"))
            result = check_for_updates(str(data.get("version") or ""), str(data.get("version_info_url") or ""))
        except Exception as exc:
            result = UpdateResult(status="error", error=str(exc))
        self.finished.emit(result)


class UpdateInstallWorker(QObject):
    progress = Signal(int)
    failed = Signal(str)
    finished = Signal(object)

    def __init__(self, package_url: str) -> None:
        super().__init__()
        self.package_url = package_url

    @Slot()
    def run(self) -> None:
        try:
            helper_path = prepare_update_install(
                package_url=self.package_url,
                target_root=project_root(),
                restart_command=self._restart_command(),
                progress_callback=self.progress.emit,
            )
        except Exception as exc:
            self.failed.emit(str(exc))
            return
        self.finished.emit(helper_path)

    def _restart_command(self) -> list[str]:
        if getattr(sys, "frozen", False):
            return [sys.executable]
        return [sys.executable, str(project_root() / "app.py")]


@dataclass(frozen=True)
class LauncherModule:
    title: str
    subtitle: str
    description: str
    icon: str
    card: str
    target: str
    button: str
    accent: str
    accent_text: str


class LauncherStage(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.background = QPixmap(str(runtime_file("assets/background.png")))
        if self.background.isNull():
            self.background = QPixmap(str(start_menu_asset("background_clean.png")))

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        painter.fillRect(self.rect(), QColor("#020A14"))
        if not self.background.isNull():
            scaled = self.background.scaled(
                self.size(),
                Qt.AspectRatioMode.KeepAspectRatioByExpanding,
                Qt.TransformationMode.SmoothTransformation,
            )
            painter.setOpacity(0.72)
            painter.drawPixmap(
                (self.width() - scaled.width()) // 2,
                (self.height() - scaled.height()) // 2,
                scaled,
            )
            painter.setOpacity(1.0)
        painter.fillRect(self.rect(), QColor(0, 0, 0, 82))
        super().paintEvent(event)


class LauncherModulesBackdrop(QFrame):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("launcherModulesBackdrop")

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        path = QPainterPath()
        path.addRoundedRect(QRectF(self.rect()), 28, 28)
        painter.setClipPath(path)
        painter.fillRect(self.rect(), QColor(2, 10, 20, 70))
        super().paintEvent(event)


class LauncherModuleCard(QFrame):
    clicked = Signal(str)

    def __init__(self, module: LauncherModule, parent=None) -> None:
        super().__init__(parent)
        self.module = module
        self.setObjectName("launcherModuleCard")
        self.setProperty("accentRole", module.accent)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumWidth(300)
        self.setMaximumWidth(360)
        self.setMinimumHeight(360)
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)
        self.card_background = QPixmap(str(start_menu_asset(module.card)))

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(6)

        visual = QLabel()
        visual.setAlignment(Qt.AlignmentFlag.AlignCenter)
        visual.setObjectName("launcherVisual")
        visual.setPixmap(QPixmap(str(start_menu_asset(module.icon))).scaled(226, 128, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        visual.setMinimumHeight(134)
        visual.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        glow = QGraphicsDropShadowEffect(visual)
        glow.setBlurRadius(32)
        glow.setOffset(0, 8)
        glow.setColor(QColor(self._effect_color(module.accent)))
        visual.setGraphicsEffect(glow)
        layout.addWidget(visual)

        title = QLabel(module.title)
        title.setProperty("launcherCardTitle", True)
        title.setWordWrap(True)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setMinimumWidth(0)
        title.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        subtitle = QLabel(module.subtitle)
        subtitle.setProperty("launcherCardSubtitle", True)
        subtitle.setProperty("accentRole", module.accent_text)
        subtitle.setWordWrap(True)
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setMinimumWidth(0)
        subtitle.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        description = QLabel(module.description)
        description.setProperty("launcherCardBody", True)
        description.setWordWrap(True)
        description.setAlignment(Qt.AlignmentFlag.AlignCenter)
        description.setMinimumWidth(0)
        description.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addWidget(description)
        layout.addStretch(1)

        button = QPushButton(module.button)
        button.setProperty("launcherAction", True)
        button.setProperty("accentRole", module.accent)
        button.setMinimumWidth(0)
        button.setMinimumHeight(42)
        button.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        button.clicked.connect(lambda _checked=False: self.clicked.emit(module.target))
        layout.addWidget(button)

    def _effect_color(self, accent: str) -> str:
        return {
            "blue": "#168CFF",
            "green": "#20E28A",
            "orange": "#FFB000",
        }.get(accent, "#20C5FF")

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        if not self.card_background.isNull():
            painter.drawPixmap(self.rect(), self.card_background)
        super().paintEvent(event)

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit(self.module.target)
        super().mouseReleaseEvent(event)


class HomePage(QScrollArea):
    navigate = Signal(str)

    MODULES = (
        LauncherModule(
            "WGA",
            "Управление и автоматизация\nна Windows",
            "Активиране на Windows и Office,\nинсталиране на пакети, драйвери,\nнастройки и полезни инструменти.",
            "wga_visual.png",
            "card_wga.png",
            "wga_dashboard",
            "Отвори WGA",
            "blue",
            "blue",
        ),
        LauncherModule(
            "Network Monitoring",
            "Мрежов мониторинг и управление",
            "Проверка на устройства в мрежата,\nIP адреси, активни клиенти,\nпринтери и споделени ресурси.",
            "network_visual.png",
            "card_network.png",
            "module::network",
            "Отвори Network Monitoring",
            "green",
            "green",
        ),
        LauncherModule(
            "Оптимизация",
            "Почистване, настройка и ускоряване",
            "Почистване на системата, управление\nна автозареждането, оптимизация\nна производителността и поверителността.",
            "optimization_visual.png",
            "card_optimization.png",
            "module::optimization",
            "Отвори Оптимизация",
            "orange",
            "orange",
        ),
    )

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setObjectName("launcherScroll")

        host = LauncherStage()
        host.setObjectName("launcherRoot")
        host.setMinimumWidth(0)
        self.setWidget(host)

        layout = QVBoxLayout(host)
        layout.setContentsMargins(22, 4, 22, 18)
        layout.setSpacing(9)

        layout.addWidget(self._window_controls())
        layout.addWidget(self._brand())
        layout.addWidget(self._module_list())
        layout.addStretch(1)
        layout.addWidget(self._footer())
        self._update_thread: QThread | None = None
        self._update_worker: UpdateCheckWorker | None = None
        self._install_thread: QThread | None = None
        self._install_worker: UpdateInstallWorker | None = None
        self._available_update: UpdateResult | None = None
        QTimer.singleShot(450, self._check_updates)

    def _window_controls(self) -> QWidget:
        row = QWidget()
        row.setObjectName("launcherChrome")
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addStretch(1)
        for label, target, role in (
            ("-", "app_minimize", "minimize"),
            ("□", "app_toggle_maximize", "maximize"),
            ("X", "app_exit", "close"),
        ):
            button = QPushButton(label)
            button.setProperty("windowControl", role)
            button.setFixedSize(38, 30)
            button.clicked.connect(lambda _checked=False, current=target: self.navigate.emit(current))
            layout.addWidget(button)
        return row

    def _brand(self) -> QFrame:
        panel = QFrame()
        panel.setObjectName("launcherBrand")
        layout = QHBoxLayout(panel)
        layout.setContentsMargins(0, 4, 0, 4)
        layout.setSpacing(18)
        layout.addStretch(1)

        logo = QLabel()
        logo.setObjectName("launcherHeroLogo")
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_pixmap = QPixmap(str(start_menu_asset("Лъскаво синьо 3D лого WGA.png")))
        if logo_pixmap.isNull():
            logo_pixmap = QPixmap(str(start_menu_asset("wga_shield_logo.png")))
        logo.setPixmap(
            logo_pixmap.scaled(
                460,
                150,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        )
        logo.setMinimumSize(470, 132)
        logo.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        logo_glow = QGraphicsDropShadowEffect(logo)
        logo_glow.setBlurRadius(34)
        logo_glow.setOffset(0, 0)
        logo_glow.setColor(QColor("#20C5FF"))
        logo.setGraphicsEffect(logo_glow)
        layout.addWidget(logo)

        layout.addSpacing(38)
        slogan = QLabel("Едно приложение.\nПовече възможности.\nЗа по-добра система.")
        slogan.setProperty("launcherSlogan", True)
        slogan.setMinimumWidth(0)
        slogan.setMaximumWidth(360)
        slogan.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        slogan.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        layout.addWidget(slogan)
        layout.addStretch(1)
        return panel

    def _module_list(self) -> QWidget:
        panel = LauncherModulesBackdrop()
        layout = QHBoxLayout(panel)
        layout.setContentsMargins(26, 18, 26, 18)
        layout.setSpacing(16)
        layout.addStretch(1)
        for module in self.MODULES:
            card = LauncherModuleCard(module)
            card.clicked.connect(self.navigate.emit)
            layout.addWidget(card)
        layout.addStretch(1)
        return panel

    def _footer(self) -> QFrame:
        footer = QFrame()
        footer.setObjectName("launcherFooter")
        layout = QHBoxLayout(footer)
        layout.setContentsMargins(18, 12, 18, 12)
        status = QLabel(f"⚙  Версия: {current_version()}")
        status.setProperty("launcherMuted", True)
        hint = QLabel("WGA  |  Your system. Our care.")
        hint.setProperty("launcherMuted", True)
        hint.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        layout.addWidget(status)
        layout.addSpacing(28)
        self.update_status = QLabel("Проверка за актуализация...")
        self.update_status.setProperty("launcherUpdateStatus", True)
        self.update_status.setMinimumWidth(0)
        layout.addWidget(self.update_status, 1)
        self.update_button = QPushButton("Провери за нова версия")
        self.update_button.setProperty("launcherUpdateButton", True)
        self.update_button.clicked.connect(self._update_button_clicked)
        layout.addWidget(self.update_button)
        layout.addStretch(1)
        layout.addWidget(hint)
        return footer

    def _update_button_clicked(self) -> None:
        if self._available_update and self._available_update.status == "update_available":
            self._show_update_prompt()
            return
        self._check_updates()

    def _check_updates(self) -> None:
        if self._update_thread is not None and self._update_thread.isRunning():
            return
        self._available_update = None
        self.update_button.setText("Провери за нова версия")
        self.update_status.setText("Проверка за актуализация...")
        self.update_status.setProperty("state", "checking")
        self._refresh_update_style()
        self.update_button.setEnabled(False)
        self._update_thread = QThread(self)
        self._update_worker = UpdateCheckWorker()
        self._update_worker.moveToThread(self._update_thread)
        self._update_thread.started.connect(self._update_worker.run)
        self._update_worker.finished.connect(self._apply_update_result)
        self._update_worker.finished.connect(self._update_thread.quit)
        self._update_worker.finished.connect(self._update_worker.deleteLater)
        self._update_thread.finished.connect(self._update_thread.deleteLater)
        self._update_thread.finished.connect(self._clear_update_thread)
        self._update_thread.start()

    def _apply_update_result(self, result: UpdateResult) -> None:
        self.update_button.setEnabled(True)
        version = current_version()
        if result.status == "up_to_date":
            self.update_status.setText(f"Програмата е актуална. Текуща версия: v{version}.")
            self.update_status.setProperty("state", "ok")
            self.update_button.setText("Провери за нова версия")
        elif result.status == "update_available":
            self.update_status.setText(f"Налична е нова версия: v{result.latest_version}.")
            self.update_status.setProperty("state", "available")
            self.update_button.setText("Изтегли и инсталирай")
            self._available_update = result
            QTimer.singleShot(250, self._show_update_prompt)
        elif result.status == "not_configured":
            self.update_status.setText("Проверката за актуализации не е настроена.")
            self.update_status.setProperty("state", "warning")
            self.update_button.setText("Провери за нова версия")
        else:
            if self._looks_like_no_internet(result.error):
                self.update_status.setText("В момента няма връзка с интернет.")
            else:
                self.update_status.setText(f"Проверката не успя: {result.error or 'неизвестна грешка'}.")
            self.update_status.setProperty("state", "error")
            self.update_button.setText("Провери отново")
        self._refresh_update_style()

    def _show_update_prompt(self) -> None:
        result = self._available_update
        if not result or result.status != "update_available":
            return
        if self._install_thread is not None and self._install_thread.isRunning():
            return
        if not (result.package_url or result.download_url):
            modern_message(self, "Липсва update пакет", "Налична е версия, но няма зададен адрес за изтегляне.")
            return
        if not modern_update_prompt(self, result.latest_version, result.notes, result.changelog):
            return
        if not result.package_url:
            modern_message(self, "Само ръчно изтегляне", "За тази версия няма автоматична актуализация. Отвори страницата за изтегляне на wsga.eu.")
            return
        self._install_update(result.package_url)

    def _install_update(self, package_url: str) -> None:
        if self._install_thread is not None and self._install_thread.isRunning():
            return
        self.update_button.setEnabled(False)
        self.update_status.setText("Подготовка за изтегляне на актуализация...")
        self.update_status.setProperty("state", "checking")
        self._refresh_update_style()

        self._install_thread = QThread(self)
        self._install_worker = UpdateInstallWorker(package_url)
        self._install_worker.moveToThread(self._install_thread)
        self._install_thread.started.connect(self._install_worker.run)
        self._install_worker.progress.connect(self._apply_install_progress)
        self._install_worker.failed.connect(self._update_install_failed)
        self._install_worker.finished.connect(self._update_install_ready)
        self._install_worker.failed.connect(self._install_thread.quit)
        self._install_worker.finished.connect(self._install_thread.quit)
        self._install_worker.failed.connect(self._install_worker.deleteLater)
        self._install_worker.finished.connect(self._install_worker.deleteLater)
        self._install_thread.finished.connect(self._install_thread.deleteLater)
        self._install_thread.finished.connect(self._clear_install_thread)
        self._install_thread.start()

    def _apply_install_progress(self, value: int) -> None:
        self.update_status.setText(f"Изтегляне и подготовка на актуализация... {value}%")
        self.update_status.setProperty("state", "checking")
        self._refresh_update_style()

    def _update_install_failed(self, detail: str) -> None:
        self.update_button.setEnabled(True)
        self.update_button.setText("Пробвай отново")
        self.update_status.setText(f"Инсталацията не успя: {detail}")
        self.update_status.setProperty("state", "error")
        self._refresh_update_style()
        modern_message(self, "Update грешка", detail)

    def _update_install_ready(self, helper_path: object) -> None:
        self.update_status.setText("Актуализацията е готова. WGA ще се рестартира.")
        self.update_status.setProperty("state", "available")
        self._refresh_update_style()
        try:
            launch_helper_and_exit(Path(str(helper_path)))
            self.navigate.emit("app_exit")
        except Exception as exc:
            self._update_install_failed(str(exc))

    def _clear_update_thread(self) -> None:
        self._update_thread = None
        self._update_worker = None

    def _clear_install_thread(self) -> None:
        self._install_thread = None
        self._install_worker = None

    def _refresh_update_style(self) -> None:
        self.update_status.style().unpolish(self.update_status)
        self.update_status.style().polish(self.update_status)

    def _looks_like_no_internet(self, error: str) -> bool:
        text = error.lower()
        return any(
            marker in text
            for marker in (
                "timed out",
                "getaddrinfo",
                "name or service not known",
                "no address associated",
                "network is unreachable",
                "temporary failure",
                "няма връзка",
                "11001",
            )
        )


class WgaDashboardPage(QScrollArea):
    navigate = Signal(str)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        host = QWidget()
        self.setWidget(host)
        layout = QVBoxLayout(host)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        layout.addWidget(self._hero())
        layout.addLayout(self._features())
        tools_title = QLabel("Полезни инструменти")
        tools_title.setProperty("title", True)
        layout.addWidget(tools_title)
        layout.addLayout(self._tools())
        layout.addStretch(1)

    def _hero(self) -> QFrame:
        panel = QFrame()
        panel.setObjectName("heroPanel")
        panel.setMinimumHeight(178)
        panel.setMaximumHeight(178)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(22, 16, 18, 16)
        layout.setSpacing(10)
        title = QLabel("WinSys Guardian Advanced")
        title.setProperty("heroTitle", True)
        subtitle = QLabel("Пълен контрол, оптимизация и автоматизация за вашия Windows")
        subtitle.setProperty("muted", True)
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addStretch(1)
        actions = QGridLayout()
        actions.setHorizontalSpacing(10)
        actions.setVerticalSpacing(0)
        for key, title_text, subtitle_text, icon in (
            ("windows", "Windows 10/11", "Активиране", "quick/windows.png"),
            ("office", "Office 2016-2024", "Активиране", "quick/office.png"),
            ("drivers", "Драйвери", "и резервно копие", "quick/drivers.png"),
        ):
            card = QuickActionCard(str(ui_asset(icon)), title_text, subtitle_text)
            card.clicked.connect(lambda current=key: self.navigate.emit(current))
            index = actions.count()
            actions.addWidget(card, 0, index)
            actions.setColumnStretch(index, 1)
        layout.addLayout(actions)
        return panel

    def _features(self) -> QHBoxLayout:
        layout = QHBoxLayout()
        layout.setSpacing(15)
        cards = (
            ("windows", "features/windows_10_11.png", "Windows 10/11", "Активация, ключове и системни инструменти", ["Windows 10 активация", "Windows 11 активация", "Desktop и системни настройки"], "#168CFF"),
            ("office", "features/office_2016_2024.png", "Office 2016-2024", "Инсталация, активация и поддръжка", ["Office 2016/2019/2021 активация", "Office Install Center", "Repair и status проверки"], "#7657FF"),
        )
        for key, icon, title, subtitle, features, accent in cards:
            card = FeatureCard(str(ui_asset(icon)), title, subtitle, features, "Отвори", accent, lambda _=False, current=key: self.navigate.emit(current))
            layout.addWidget(card, 1)
        return layout

    def _tools(self) -> QGridLayout:
        layout = QGridLayout()
        layout.setSpacing(8)
        tools = [
            ("tools/cleaning.png", "Почистване", "Отваря наличните maintenance действия.", "system"),
            ("tools/services.png", "Управление на услуги", "Системни инструменти.", "system"),
            ("tools/startup.png", "Автостарт", "Инструменти и install център.", "tools"),
            ("tools/personalization.png", "Персонализация", "Език, клавиатури и desktop опции.", "personalization"),
            ("tools/privacy.png", "Защита и поверителност", "Административни проверки и статуси.", "privacy"),
            ("tools/drivers.png", "Драйвери", "Backup, restore, report и recovery USB.", "drivers"),
            ("tools/onedrive.png", "OneDrive", "Reset workflows за OneDrive.", "tools"),
            ("tools/network.png", "Мрежа и интернет", "Network monitoring module.", "network"),
            ("tools/system_tools.png", "Системни инструменти", "Nexus admin, health и скрити инструменти.", "advanced"),
            ("tools/recovery_usb.png", "Възстановителен USB", "Driver recovery USB workflow.", "drivers"),
        ]
        columns = 5
        for index, (icon, title, desc, key) in enumerate(tools):
            card = ToolCard(str(ui_asset(icon)), title, desc, lambda _=False, current=key: self.navigate.emit(current))
            layout.addWidget(card, index // columns, index % columns)
        return layout
