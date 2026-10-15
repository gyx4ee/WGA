from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton, QSizePolicy, QVBoxLayout, QWidget


class FeatureCard(QFrame):
    def __init__(
        self,
        icon: str,
        title: str,
        subtitle: str,
        features: list[str],
        button_text: str,
        accent: str,
        callback,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setProperty("card", True)
        self.setProperty("feature", True)
        self.setMinimumWidth(0)
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 11, 12, 11)
        layout.setSpacing(5)
        top = QHBoxLayout()
        image = QLabel()
        image.setPixmap(QPixmap(icon).scaled(34, 34, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        top.addWidget(image)
        titles = QVBoxLayout()
        title_label = QLabel(title)
        title_label.setProperty("subtitle", True)
        title_label.setWordWrap(True)
        subtitle_label = QLabel(subtitle)
        subtitle_label.setWordWrap(True)
        subtitle_label.setProperty("muted", True)
        titles.addWidget(title_label)
        titles.addWidget(subtitle_label)
        top.addLayout(titles, 1)
        layout.addLayout(top)
        for feature in features:
            row = QLabel(f"• {feature}")
            row.setProperty("muted", True)
            layout.addWidget(row)
        layout.addStretch(1)
        button = QPushButton(button_text)
        button.setProperty("accent", accent)
        button.clicked.connect(callback)
        layout.addWidget(button)


class ToolCard(QFrame):
    def __init__(self, icon: str, title: str, description: str, callback, parent=None) -> None:
        super().__init__(parent)
        self.setProperty("card", True)
        self.setProperty("tool", True)
        self.setMinimumWidth(0)
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(9, 9, 9, 9)
        layout.setSpacing(4)
        image = QLabel()
        image.setAlignment(Qt.AlignmentFlag.AlignCenter)
        image.setPixmap(QPixmap(icon).scaled(28, 28, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        layout.addWidget(image)
        title_label = QLabel(title)
        title_label.setProperty("subtitle", True)
        title_label.setWordWrap(True)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)
        body = QLabel(description)
        body.setWordWrap(True)
        body.setProperty("muted", True)
        body.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(body)
        layout.addStretch(1)
        button = QPushButton("Отвори")
        button.setProperty("secondary", True)
        button.clicked.connect(callback)
        layout.addWidget(button)


class QuickActionCard(QFrame):
    clicked = Signal()

    def __init__(self, icon: str, title: str, subtitle: str = "", parent=None) -> None:
        super().__init__(parent)
        self.setProperty("quick", True)
        self.setMinimumWidth(0)
        self.setMinimumHeight(52)
        self.setMaximumHeight(56)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(9, 6, 9, 6)
        layout.setSpacing(7)
        image = QLabel()
        image.setPixmap(QPixmap(icon).scaled(26, 26, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        layout.addWidget(image)
        text = QVBoxLayout()
        title_label = QLabel(title)
        title_label.setProperty("subtitle", True)
        title_label.setWordWrap(True)
        text.addWidget(title_label)
        if subtitle:
            sub = QLabel(subtitle)
            sub.setProperty("muted", True)
            sub.setWordWrap(True)
            text.addWidget(sub)
        layout.addLayout(text, 1)

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mouseReleaseEvent(event)
