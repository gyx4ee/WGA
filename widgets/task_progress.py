from __future__ import annotations

from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton, QProgressBar, QVBoxLayout


class TaskProgressWidget(QFrame):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setProperty("card", True)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        self.title = QLabel("Текуща задача")
        self.title.setProperty("subtitle", True)
        self.operation = QLabel("Няма активна операция")
        self.operation.setProperty("muted", True)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        footer = QHBoxLayout()
        self.elapsed = QLabel("00:00")
        self.elapsed.setProperty("muted", True)
        self.stop_button = QPushButton("Stop")
        self.stop_button.setEnabled(False)
        footer.addWidget(self.elapsed)
        footer.addStretch(1)
        footer.addWidget(self.stop_button)
        layout.addWidget(self.title)
        layout.addWidget(self.operation)
        layout.addWidget(self.progress)
        layout.addLayout(footer)
