from __future__ import annotations

import re
from dataclasses import dataclass

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import QFrame, QGridLayout, QHBoxLayout, QLabel, QPushButton, QScrollArea, QSizePolicy, QVBoxLayout, QWidget

from services.paths import ui_asset


@dataclass(frozen=True)
class SectionAction:
    title: str
    description: str
    target: str
    icon: str
    accent: str = "#168CFF"
    external: bool = False


class SectionActionCard(QFrame):
    activated = Signal(str, bool)

    def __init__(self, action: SectionAction, parent=None) -> None:
        super().__init__(parent)
        self.action = action
        self.setProperty("card", True)
        self.setProperty("sectionAction", True)
        self.setMinimumWidth(0)
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(9)
        top = QHBoxLayout()
        image = QLabel()
        image.setPixmap(QPixmap(str(ui_asset(action.icon))).scaled(36, 36, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        image.setMinimumSize(40, 40)
        top.addWidget(image)
        heading = QLabel(action.title)
        heading.setProperty("subtitle", True)
        heading.setWordWrap(True)
        top.addWidget(heading, 1)
        layout.addLayout(top)
        body = QLabel(action.description)
        body.setWordWrap(True)
        body.setProperty("muted", True)
        layout.addWidget(body)
        layout.addStretch(1)
        button = QPushButton("Отвори")
        button.setProperty("accent", action.accent)
        button.clicked.connect(lambda _checked=False: self.activated.emit(action.target, action.external))
        layout.addWidget(button)

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.MouseButton.LeftButton:
            self.activated.emit(self.action.target, self.action.external)
        super().mouseReleaseEvent(event)


class SimpleSectionPage(QScrollArea):
    navigate = Signal(str)

    def __init__(self, title: str, description: str, legacy_key: str, actions: list[SectionAction] | None = None, parent=None) -> None:
        super().__init__(parent)
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        host = QWidget()
        self.setWidget(host)
        layout = QVBoxLayout(host)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)

        intro = QFrame()
        intro.setProperty("card", True)
        intro.setProperty("sectionIntro", True)
        intro_layout = QVBoxLayout(intro)
        intro_layout.setContentsMargins(18, 16, 18, 16)
        intro_layout.setSpacing(8)
        label = QLabel(title)
        label.setProperty("title", True)
        intro_layout.addWidget(label)
        body = QLabel(description)
        body.setWordWrap(True)
        body.setProperty("muted", True)
        intro_layout.addWidget(body)
        open_module = QPushButton("Отвори раздела")
        open_module.clicked.connect(lambda: self._activate(legacy_key, False))
        intro_layout.addWidget(open_module, 0, Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(intro)

        grid = QGridLayout()
        grid.setSpacing(12)
        section_actions = actions or [
            SectionAction("Основен раздел", description, legacy_key, "tools/system_tools.png"),
        ]
        for index, action in enumerate(section_actions):
            card = SectionActionCard(action)
            card.activated.connect(self._activate)
            grid.addWidget(card, index // 3, index % 3)
        layout.addLayout(grid)
        layout.addStretch(1)

    def _activate(self, target: str, external: bool) -> None:
        self.navigate.emit(target)


class LegacyMenuPage(QScrollArea):
    navigate = Signal(str)

    def __init__(self, menu_key: str, menu_data: dict[str, object], parent_key: str | None = None, parent=None) -> None:
        super().__init__(parent)
        self.menu_key = menu_key
        self.parent_key = parent_key
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        host = QWidget()
        self.setWidget(host)
        layout = QVBoxLayout(host)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)

        intro = QFrame()
        intro.setProperty("card", True)
        intro.setProperty("sectionIntro", True)
        intro_layout = QVBoxLayout(intro)
        intro_layout.setContentsMargins(18, 16, 18, 16)
        intro_layout.setSpacing(8)
        title = QLabel(str(menu_data.get("title") or menu_key))
        title.setProperty("title", True)
        intro_layout.addWidget(title)
        subtitle = QLabel(str(menu_data.get("subtitle") or "Изберете инструмент от това меню."))
        subtitle.setWordWrap(True)
        subtitle.setProperty("muted", True)
        intro_layout.addWidget(subtitle)
        layout.addWidget(intro)

        items = [item for item in menu_data.get("items", []) if isinstance(item, dict) and not self._is_legacy_back_item(item)]
        grid = QGridLayout()
        grid.setSpacing(12)
        if not items:
            empty = QLabel("Това меню няма отделни карти в момента.")
            empty.setProperty("muted", True)
            layout.addWidget(empty)
        for index, item in enumerate(items):
            action = self._action_from_item(item)
            card = SectionActionCard(action)
            card.activated.connect(self._activate)
            grid.addWidget(card, index // 3, index % 3)
        layout.addLayout(grid)
        if self.parent_key:
            back_button = QPushButton("< Назад")
            back_button.setProperty("danger", True)
            back_button.clicked.connect(lambda _checked=False: self.navigate.emit("home"))
            layout.addWidget(back_button, 0, Qt.AlignmentFlag.AlignHCenter)
        layout.addStretch(1)

    def _is_legacy_back_item(self, item: dict[str, object]) -> bool:
        label = str(item.get("label") or "").strip().lower()
        target = str(item.get("target") or "").strip().lower()
        return label in {"return to main menu", "назад", "back"} or target == "main"

    def _action_from_item(self, item: dict[str, object]) -> SectionAction:
        kind = str(item.get("kind") or "action")
        label = str(item.get("label") or "Инструмент")
        description = str(item.get("description") or ("Отваря подменюто." if kind == "menu" else "Стартира WGA инструмента."))
        target = str(item.get("target") or self.menu_key)
        if kind == "exit":
            target = "app_exit"
            description = "Затваря WinSys Guardian Advanced."
        elif kind != "menu":
            action_id = self._slug(str(item.get("action_id") or label))
            target = f"execute::{self.menu_key}::{action_id}"
        icon = self._icon_for_item(kind, label)
        accent = "#20C5FF" if kind == "menu" else "#B92D4A" if kind == "exit" else "#168CFF"
        return SectionAction(label, description, target, icon, accent, external=False)

    def _activate(self, target: str, external: bool) -> None:
        self.navigate.emit(target)

    def _icon_for_item(self, kind: str, label: str) -> str:
        text = label.lower()
        if "office" in text:
            return "navigation/office.png"
        if "windows" in text:
            return "navigation/windows.png"
        if "driver" in text or "драй" in text:
            return "navigation/drivers.png"
        if "onedrive" in text:
            return "tools/onedrive.png"
        if "language" in text or "език" in text:
            return "tools/personalization.png"
        if "admin" in text or "nexus" in text:
            return "navigation/system.png"
        if "install" in text or "инстал" in text:
            return "tools/services.png"
        if kind == "menu":
            return "actions/open.png"
        return "actions/start.png"

    def _slug(self, value: str) -> str:
        return re.sub(r"[^a-z0-9]+", "_", value.strip().lower()).strip("_") or "action"


