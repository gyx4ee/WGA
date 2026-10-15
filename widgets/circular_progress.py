from __future__ import annotations

from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QPainter, QPen
from PySide6.QtWidgets import QWidget


class CircularProgress(QWidget):
    def __init__(self, label: str, value: int = 0, accent: str = "#20C5FF", size: int = 104, parent=None) -> None:
        super().__init__(parent)
        self.label = label
        self.value = max(0, min(100, value))
        self.accent = QColor(accent)
        self.setMinimumSize(size, size)
        self.setMaximumSize(size, size)

    def set_value(self, value: int) -> None:
        self.value = max(0, min(100, value))
        self.update()

    def paintEvent(self, _event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        side = min(self.width(), self.height()) - 16
        rect = QRectF((self.width() - side) / 2, 8, side, side)
        pen_width = max(6, int(side * 0.1))
        pen = QPen(QColor("#173555"), pen_width)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        painter.setPen(pen)
        painter.drawArc(rect, 0, 360 * 16)
        pen.setColor(self.accent)
        painter.setPen(pen)
        painter.drawArc(rect, 90 * 16, -int(360 * 16 * self.value / 100))
        painter.setPen(QColor("#F5F8FF"))
        font = painter.font()
        font.setPointSize(max(10, int(side * 0.18)))
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, f"{self.value}%")
        font.setPointSize(max(7, int(side * 0.1)))
        font.setBold(False)
        painter.setFont(font)
        painter.setPen(QColor("#AFC5DF"))
        painter.drawText(rect.adjusted(0, side * 0.28, 0, side * 0.28), Qt.AlignmentFlag.AlignCenter, self.label)
