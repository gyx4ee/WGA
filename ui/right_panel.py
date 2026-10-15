from __future__ import annotations

import subprocess

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import QFrame, QGridLayout, QHBoxLayout, QLabel, QProgressBar, QPushButton, QScrollArea, QVBoxLayout, QWidget

from services.paths import ui_asset
from services.system_info import static_system_rows, usage_snapshot
from ui.dialogs import modern_confirm
from widgets.circular_progress import CircularProgress
from widgets.task_progress import TaskProgressWidget


class RightPanel(QFrame):
    navigate = Signal(str)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setFixedWidth(316)
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_body = QWidget()
        scroll.setWidget(scroll_body)
        outer_layout.addWidget(scroll)

        layout = QVBoxLayout(scroll_body)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        status = QFrame()
        status.setProperty("card", True)
        status.setMinimumHeight(390)
        status.setMaximumHeight(390)
        status_layout = QVBoxLayout(status)
        status_layout.setContentsMargins(12, 12, 12, 12)
        status_layout.setSpacing(8)
        status_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        title = QLabel("Състояние на системата")
        title.setProperty("subtitle", True)
        status_layout.addWidget(title)
        self.cpu, self.cpu_detail, self.cpu_circle = self._metric_row(status_layout, "CPU", "Процесор (CPU)", "Зареждане...", "#20C5FF")
        self.ram, self.ram_detail, self.ram_circle = self._metric_row(status_layout, "RAM", "Памет (RAM)", "Зареждане...", "#7657FF")
        self.disk, self.disk_detail, self.disk_circle = self._metric_row(status_layout, "Disk", "Всички дискове", "Зареждане...", "#20E28A")
        self.gpu, self.gpu_detail, self.gpu_circle = self._metric_row(status_layout, "GPU", "Графика (GPU)", "Зареждане...", "#8A5CFF")
        layout.addWidget(status)

        info = QFrame()
        info.setProperty("card", True)
        info.setMinimumHeight(244)
        info.setMaximumHeight(292)
        info_layout = QVBoxLayout(info)
        info_layout.setContentsMargins(12, 12, 12, 12)
        info_layout.setSpacing(8)
        info_title = QLabel("Системна информация")
        info_title.setProperty("subtitle", True)
        info_layout.addWidget(info_title)
        info_scroll = QScrollArea()
        info_scroll.setWidgetResizable(True)
        info_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        info_body = QWidget()
        info_body_layout = QVBoxLayout(info_body)
        info_body_layout.setContentsMargins(0, 0, 0, 0)
        info_body_layout.setSpacing(7)
        for label, value in static_system_rows():
            info_body_layout.addWidget(self._info_tile(label, value))
        info_body_layout.addStretch(1)
        info_scroll.setWidget(info_body)
        info_layout.addWidget(info_scroll)
        layout.addWidget(info)

        quick = QFrame()
        quick.setProperty("card", True)
        quick.setMaximumHeight(145)
        quick_layout = QVBoxLayout(quick)
        quick_title = QLabel("Бързи действия")
        quick_title.setProperty("subtitle", True)
        quick_layout.addWidget(quick_title)
        grid = QGridLayout()
        grid.setSpacing(8)
        action_index = 0
        for text, callback in (
            ("Рестарт", self._confirm_restart),
            ("Изключване", self._confirm_shutdown),
            ("GOD MODE", lambda: self.navigate.emit("hidden_menu")),
            ("PC отчет", lambda: self.navigate.emit("driver_backup")),
        ):
            button = QPushButton(text)
            button.clicked.connect(callback)
            grid.addWidget(button, action_index // 2, action_index % 2)
            action_index += 1
        quick_layout.addLayout(grid)
        layout.addWidget(quick)
        task_widget = TaskProgressWidget()
        task_widget.setMaximumHeight(145)
        layout.addWidget(task_widget)
        layout.addStretch(1)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.refresh)
        self.timer.start(2000)
        self.refresh()

    def _metric_row(self, parent_layout: QVBoxLayout, short_label: str, title: str, detail: str, accent: str) -> tuple[QProgressBar, QLabel, CircularProgress]:
        row = QFrame()
        row.setProperty("metric", True)
        row.setMinimumHeight(86)
        row.setMaximumHeight(86)
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(10, 8, 10, 8)
        row_layout.setSpacing(10)
        circle = CircularProgress(short_label, accent=accent, size=56)
        row_layout.addWidget(circle, 0, Qt.AlignmentFlag.AlignVCenter)
        text_layout = QVBoxLayout()
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(5)
        label = QLabel(title)
        label.setProperty("metricTitle", True)
        label.setWordWrap(True)
        detail_label = QLabel(detail)
        detail_label.setProperty("metricDetail", True)
        detail_label.setWordWrap(True)
        progress = QProgressBar()
        progress.setRange(0, 100)
        progress.setTextVisible(True)
        progress.setFormat(f"{short_label}: %p%")
        progress.setProperty("metricBar", True)
        progress.setStyleSheet(f"QProgressBar::chunk {{ background: {accent}; border-radius: 6px; }}")
        text_layout.addWidget(label)
        text_layout.addWidget(detail_label)
        text_layout.addWidget(progress)
        row_layout.addLayout(text_layout, 1)
        parent_layout.addWidget(row)
        return progress, detail_label, circle

    def _info_tile(self, label_text: str, value: str) -> QFrame:
        tile = QFrame()
        tile.setProperty("infoTile", True)
        tile.setMinimumHeight(68)
        tile_layout = QHBoxLayout(tile)
        tile_layout.setContentsMargins(8, 7, 8, 7)
        tile_layout.setSpacing(7)
        icon = QLabel()
        icon.setMinimumSize(28, 28)
        icon.setMaximumSize(28, 28)
        pixmap = QPixmap(str(ui_asset(self._info_icon(label_text))))
        if not pixmap.isNull():
            icon.setPixmap(pixmap.scaled(26, 26, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        tile_layout.addWidget(icon, 0, Qt.AlignmentFlag.AlignTop)
        text = QVBoxLayout()
        text.setContentsMargins(0, 0, 0, 0)
        text.setSpacing(2)
        label = QLabel(label_text)
        label.setProperty("infoTitle", True)
        label.setWordWrap(True)
        detail = QLabel(self._shorten_info(value))
        detail.setProperty("infoDetail", True)
        detail.setWordWrap(True)
        text.addWidget(label)
        text.addWidget(detail)
        tile_layout.addLayout(text, 1)
        return tile

    def _info_icon(self, label_text: str) -> str:
        label = label_text.lower()
        if "windows" in label or "версия" in label or "build" in label:
            return "system/windows_info.png"
        if "cpu" in label:
            return "system/cpu.png"
        if "ram" in label:
            return "system/ram.png"
        if "gpu" in label:
            return "system/gpu.png"
        if "motherboard" in label:
            return "system/motherboard.png"
        if "диск" in label:
            return "system/disk.png"
        return "status/ready.png"

    def _shorten_info(self, value: str) -> str:
        clean = value.replace("AuthenticAMD", "AMD").replace("GenuineIntel", "Intel")
        return clean if len(clean) <= 44 else clean[:41].rstrip() + "..."

    def refresh(self) -> None:
        snapshot = usage_snapshot()
        self.cpu.setValue(snapshot.cpu)
        self.ram.setValue(snapshot.ram)
        self.disk.setValue(snapshot.disk)
        self.gpu.setValue(0)
        self.cpu_circle.set_value(snapshot.cpu)
        self.ram_circle.set_value(snapshot.ram)
        self.disk_circle.set_value(snapshot.disk)
        self.gpu_circle.set_value(0)
        self.cpu_detail.setText(f"{snapshot.cpu}% използване")
        self.ram_detail.setText(f"{snapshot.ram}% използване")
        self.disk_detail.setText(snapshot.disk_detail)
        self.gpu_detail.setText(snapshot.gpu)

    def _confirm_restart(self) -> None:
        if modern_confirm(self, "Рестарт", "Сигурни ли сте, че искате да рестартирате компютъра?"):
            subprocess.Popen(["shutdown", "/r", "/t", "5"], close_fds=True)

    def _confirm_shutdown(self) -> None:
        if modern_confirm(self, "Изключване", "Сигурни ли сте, че искате да изключите компютъра?"):
            subprocess.Popen(["shutdown", "/s", "/t", "5"], close_fds=True)
