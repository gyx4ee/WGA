from __future__ import annotations

from PySide6.QtCore import Signal, Qt
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

from services.paths import ui_asset
from services.version import current_version


class Sidebar(QFrame):
    section_selected = Signal(str)
    exit_requested = Signal()

    NAV_ITEMS = [
        ("home", "Начало", "navigation/home.png"),
        ("windows", "Windows", "navigation/windows.png"),
        ("office", "Office", "navigation/office.png"),
        ("software", "Софтуер и пакети", "navigation/tools.png"),
        ("drivers", "Драйвери", "navigation/drivers.png"),
        ("network", "Мрежа и интернет", "navigation/network.png"),
        ("system", "Система", "navigation/system.png"),
        ("tools", "Инструменти", "navigation/tools.png"),
        ("personalization", "Персонализация", "navigation/personalization.png"),
        ("privacy", "Защита и поверителност", "navigation/privacy.png"),
        ("advanced", "Разширени", "navigation/advanced.png"),
    ]

    BOTTOM_ITEMS = [
        ("settings", "Настройки", "navigation/settings.png"),
        ("about", "Информация", "navigation/information.png"),
    ]

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("sidebar")
        self.setFixedWidth(274)
        self.buttons: dict[str, QPushButton] = {}
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(8)
        brand = QWidget()
        brand_layout = QHBoxLayout(brand)
        brand_layout.setContentsMargins(0, 4, 0, 8)
        brand_layout.setSpacing(12)
        symbol = QLabel()
        symbol.setAlignment(Qt.AlignmentFlag.AlignCenter)
        symbol.setPixmap(QPixmap(str(ui_asset("logo/wga_symbol.png"))).scaled(94, 94, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        symbol.setMinimumSize(94, 94)
        symbol.setMaximumSize(94, 94)
        brand_layout.addWidget(symbol)
        text_layout = QVBoxLayout()
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(2)
        title = QLabel("WGA")
        title.setProperty("brandTitle", True)
        subtitle = QLabel(f"WinSys Guardian\nAdvanced\nv{current_version()}")
        subtitle.setProperty("brandSubtitle", True)
        text_layout.addStretch(1)
        text_layout.addWidget(title)
        text_layout.addWidget(subtitle)
        text_layout.addStretch(1)
        brand_layout.addLayout(text_layout, 1)
        layout.addWidget(brand)
        layout.addSpacing(10)
        for key, label, icon_path in self.NAV_ITEMS:
            layout.addWidget(self._button(key, label, icon_path))
        layout.addStretch(1)
        for key, label, icon_path in self.BOTTOM_ITEMS:
            layout.addWidget(self._button(key, label, icon_path))
        exit_button = QPushButton("Изход")
        exit_button.setProperty("nav", True)
        exit_button.setIcon(QIcon(str(ui_asset("actions/stop.png"))))
        exit_button.setIconSize(exit_button.iconSize() * 1.25)
        exit_button.clicked.connect(self.exit_requested.emit)
        layout.addWidget(exit_button)
        self.set_selected("home")

    def _button(self, key: str, label: str, icon_path: str) -> QPushButton:
        button = QPushButton(label)
        button.setProperty("nav", True)
        button.setIcon(QIcon(str(ui_asset(icon_path))))
        button.setIconSize(button.iconSize() * 1.25)
        button.clicked.connect(lambda _checked=False, current=key: self.section_selected.emit(current))
        self.buttons[key] = button
        return button

    def set_selected(self, key: str) -> None:
        for item_key, button in self.buttons.items():
            button.setProperty("selected", item_key == key)
            button.style().unpolish(button)
            button.style().polish(button)
