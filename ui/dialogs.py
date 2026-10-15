from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import QComboBox, QDialog, QHBoxLayout, QLabel, QLineEdit, QPushButton, QTextEdit, QVBoxLayout

from services.paths import ui_asset


class ModernDialog(QDialog):
    def __init__(self, title: str, message: str, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint, True)
        self.setModal(True)
        self.setMinimumWidth(520)
        self.setObjectName("modernDialog")
        self.result_value: str | None = None

        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(18, 16, 18, 16)
        self.layout.setSpacing(12)

        header = QHBoxLayout()
        logo = QLabel()
        logo.setPixmap(QPixmap(str(ui_asset("logo/wga_symbol.png"))).scaled(44, 44, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        header.addWidget(logo)
        text = QVBoxLayout()
        title_label = QLabel(title)
        title_label.setProperty("subtitle", True)
        message_label = QLabel(message)
        message_label.setWordWrap(True)
        message_label.setProperty("muted", True)
        text.addWidget(title_label)
        text.addWidget(message_label)
        header.addLayout(text, 1)
        close_button = QPushButton("X")
        close_button.setProperty("secondary", True)
        close_button.setFixedWidth(42)
        close_button.clicked.connect(self.reject)
        header.addWidget(close_button)
        self.layout.addLayout(header)

    def _button_row(self, primary_text: str, secondary_text: str = "Отказ") -> QHBoxLayout:
        row = QHBoxLayout()
        row.addStretch(1)
        secondary = QPushButton(secondary_text)
        secondary.setProperty("secondary", True)
        secondary.clicked.connect(self.reject)
        primary = QPushButton(primary_text)
        primary.clicked.connect(self.accept)
        row.addWidget(secondary)
        row.addWidget(primary)
        return row


class ModernTextInputDialog(ModernDialog):
    def __init__(self, title: str, message: str, value: str = "", parent=None, password: bool = False) -> None:
        super().__init__(title, message, parent)
        self.input = QLineEdit()
        self.input.setText(value)
        if password:
            self.input.setEchoMode(QLineEdit.EchoMode.Password)
        self.input.setMinimumHeight(38)
        self.layout.addWidget(self.input)
        self.layout.addLayout(self._button_row("Запази"))

    def accept(self) -> None:
        self.result_value = self.input.text()
        super().accept()


class ModernChoiceDialog(ModernDialog):
    def __init__(self, title: str, message: str, choices: list[str], parent=None) -> None:
        super().__init__(title, message, parent)
        self.combo = QComboBox()
        self.combo.addItems(choices)
        self.combo.setMinimumHeight(38)
        self.layout.addWidget(self.combo)
        self.layout.addLayout(self._button_row("Избери"))

    def accept(self) -> None:
        self.result_value = self.combo.currentText()
        super().accept()


def modern_text_input(parent, title: str, message: str, value: str = "", password: bool = False) -> tuple[str, bool]:
    dialog = ModernTextInputDialog(title, message, value, parent, password=password)
    accepted = dialog.exec() == QDialog.DialogCode.Accepted
    return dialog.result_value or "", accepted


def modern_choice(parent, title: str, message: str, choices: list[str]) -> tuple[str, bool]:
    dialog = ModernChoiceDialog(title, message, choices, parent)
    accepted = dialog.exec() == QDialog.DialogCode.Accepted
    return dialog.result_value or "", accepted


def modern_message(parent, title: str, message: str, button_text: str = "OK") -> None:
    dialog = ModernDialog(title, message, parent)
    dialog.layout.addLayout(dialog._button_row(button_text, "Затвори"))
    dialog.exec()


def modern_details_message(parent, title: str, message: str, details: str, button_text: str = "Затвори") -> None:
    dialog = ModernDialog(title, message, parent)
    dialog.resize(760, 620)
    details_box = QTextEdit()
    details_box.setReadOnly(True)
    details_box.setPlainText(details)
    details_box.setMinimumHeight(420)
    dialog.layout.addWidget(details_box)
    dialog.layout.addLayout(dialog._button_row(button_text, "Затвори"))
    dialog.exec()


def modern_confirm(parent, title: str, message: str) -> bool:
    dialog = ModernDialog(title, message, parent)
    dialog.layout.addLayout(dialog._button_row("Да", "Не"))
    return dialog.exec() == QDialog.DialogCode.Accepted


def modern_update_prompt(parent, latest_version: str, notes: str, changelog: tuple[str, ...]) -> bool:
    dialog = ModernDialog(
        "Налична е нова версия",
        f"WinSys Guardian Advanced v{latest_version} е готов за изтегляне и инсталация.",
        parent,
    )
    dialog.resize(720, 560)

    if notes.strip():
        notes_label = QLabel(notes.strip())
        notes_label.setWordWrap(True)
        notes_label.setProperty("muted", True)
        dialog.layout.addWidget(notes_label)

    details = QTextEdit()
    details.setReadOnly(True)
    details.setMinimumHeight(300)
    details.setPlainText("\n".join(changelog[:8]) if changelog else "Няма добавена история за тази версия.")
    dialog.layout.addWidget(details)
    dialog.layout.addLayout(dialog._button_row("Изтегли и инсталирай", "По-късно"))
    return dialog.exec() == QDialog.DialogCode.Accepted
