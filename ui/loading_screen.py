from __future__ import annotations

from PySide6.QtCore import QEasingCurve, QElapsedTimer, QPropertyAnimation, QRectF, Qt, QThread, QTimer, Signal
from PySide6.QtGui import QColor, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QGraphicsOpacityEffect,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from services.paths import loading_asset
from services.startup_loader import STARTUP_STEPS, StartupStep, StartupWorker


class GoldProgressBar(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._value = 0
        self.setMinimumHeight(34)
        self.setMaximumHeight(34)

    def value(self) -> int:
        return self._value

    def set_value(self, value: int) -> None:
        self._value = max(0, min(100, value))
        self.update()

    def paintEvent(self, event) -> None:  # noqa: N802
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(0.5, 6.5, self.width() - 1, 21)
        painter.setPen(QPen(QColor(230, 219, 198, 190), 1.2))
        painter.setBrush(QColor(13, 25, 35, 190))
        painter.drawRoundedRect(rect, 10.5, 10.5)
        fill_width = max(0.0, rect.width() * (self._value / 100))
        if fill_width > 0:
            fill = QRectF(rect.x(), rect.y(), fill_width, rect.height())
            gradient = QColor(255, 199, 87)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(gradient)
            painter.drawRoundedRect(fill, 10.5, 10.5)
            painter.setBrush(QColor(255, 241, 176, 170))
            painter.drawRoundedRect(QRectF(fill.x(), fill.y() + 2, fill.width(), 5), 6, 6)


class LoadingStepRow(QWidget):
    def __init__(self, step: StartupStep, parent=None) -> None:
        super().__init__(parent)
        self.step = step
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 2, 0, 2)
        layout.setSpacing(10)
        self.icon = QLabel("...")
        self.icon.setFixedWidth(28)
        self.icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label = QLabel(step.label)
        self.label.setWordWrap(True)
        layout.addWidget(self.icon)
        layout.addWidget(self.label, 1)
        self.set_status("pending")

    def set_status(self, status: str) -> None:
        palette = {
            "pending": ("...", "#717D8C", 400),
            "active": ("◌", "#F8B94A", 800),
            "done": ("✓", "#61D98B", 700),
            "warning": ("!", "#F8B94A", 700),
            "error": ("X", "#FF5C7A", 800),
        }
        text, color, weight = palette[status]
        self.icon.setText(text)
        self.icon.setStyleSheet(f"color: {color}; font-size: 18px; font-weight: 900; background: transparent;")
        self.label.setStyleSheet(f"color: {color}; font-size: 15px; font-weight: {weight}; background: transparent;")


class LoadingScreen(QWidget):
    completed = Signal(object)
    failed = Signal(str)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("WinSys Guardian Advanced")
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint, True)
        self.resize(1100, 680)
        self.setMinimumSize(900, 560)
        self._minimum_visible_ms = 3200
        self._visible_timer = QElapsedTimer()
        self._visible_timer.start()
        self._target_progress = 0
        self._current_progress = 0
        self._wave_shift = 0
        self._warnings: list[str] = []
        self._thread: QThread | None = None
        self._worker: StartupWorker | None = None
        self._background = QPixmap(str(loading_asset("background_clean.png")))
        self._planet = QPixmap(str(loading_asset("planet_overlay.png")))
        self._waves_1 = QPixmap(str(loading_asset("gold_waves_1.png")))
        self._waves_2 = QPixmap(str(loading_asset("gold_waves_2.png")))
        self._line = QPixmap(str(loading_asset("gold_light_line.png")))
        self._pending_window: object | None = None
        self._build_ui()
        self._progress_timer = QTimer(self)
        self._progress_timer.timeout.connect(self._animate_progress)
        self._progress_timer.start(16)
        self._wave_timer = QTimer(self)
        self._wave_timer.timeout.connect(self._animate_waves)
        self._wave_timer.start(45)

    def _build_ui(self) -> None:
        self.opacity = QGraphicsOpacityEffect(self)
        self.opacity.setOpacity(0.0)
        self.setGraphicsEffect(self.opacity)
        fade = QPropertyAnimation(self.opacity, b"opacity", self)
        fade.setDuration(420)
        fade.setStartValue(0.0)
        fade.setEndValue(1.0)
        fade.setEasingCurve(QEasingCurve.Type.OutCubic)
        fade.start(QPropertyAnimation.DeletionPolicy.DeleteWhenStopped)

        root = QVBoxLayout(self)
        root.setContentsMargins(42, 30, 42, 30)
        root.setSpacing(0)
        top = QHBoxLayout()
        version = QLabel("Версия: WGA")
        version.setStyleSheet("color: #B7C6D8; font-size: 16px; background: transparent;")
        top.addWidget(version)
        top.addStretch(1)
        slogan = QLabel("По-добра система.\nПо-голяма сигурност.\nПовече възможности.")
        slogan.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        slogan.setStyleSheet("color: #FAD37A; font-size: 24px; font-style: italic; background: transparent;")
        top.addWidget(slogan)
        top.addSpacing(28)
        self.close_button = QPushButton("X")
        self.close_button.setFixedSize(38, 30)
        self.close_button.setStyleSheet("QPushButton { background: transparent; color: #D7E2F0; border: 0; font-weight: 800; } QPushButton:hover { background: #7A2238; border-radius: 8px; }")
        self.close_button.clicked.connect(self.close)
        top.addWidget(self.close_button)
        root.addLayout(top)
        root.addStretch(1)

        center = QVBoxLayout()
        center.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.logo = QLabel()
        self.logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_pixmap = QPixmap(str(loading_asset("logo.png")))
        uses_image_logo = not logo_pixmap.isNull()
        if uses_image_logo:
            self.logo.setPixmap(
                logo_pixmap.scaled(
                    560,
                    188,
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation,
                )
            )
        else:
            self.logo.setText("WGA")
            self.logo.setStyleSheet("color: #FFD77C; font-size: 118px; font-weight: 900; letter-spacing: 0; background: transparent;")
        self.logo.setMinimumHeight(168)
        self.logo.setStyleSheet(self.logo.styleSheet() + "background: transparent;")
        self.title = QLabel("WinSys Guardian Advanced")
        self.title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title.setStyleSheet("color: #FFF0C0; font-size: 42px; font-weight: 800; background: transparent;")
        self.subtitle = QLabel("WINDOWS SYSTEM MANAGEMENT & AUTOMATION")
        self.subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.subtitle.setStyleSheet("color: #F4F7FF; font-size: 16px; background: transparent;")
        tag_line = QLabel("PEOPLE  •  IDEAS  •  OPPORTUNITIES")
        tag_line.setAlignment(Qt.AlignmentFlag.AlignCenter)
        tag_line.setStyleSheet("color: #D8D6D0; font-size: 15px; background: transparent;")
        center.addWidget(self.logo)
        if not uses_image_logo:
            center.addWidget(self.title)
            center.addWidget(self.subtitle)
            center.addWidget(tag_line)
        root.addLayout(center)

        progress_row = QHBoxLayout()
        progress_row.setContentsMargins(220, 18, 220, 8)
        self.progress_bar = GoldProgressBar()
        self.percent = QLabel("0%")
        self.percent.setFixedWidth(86)
        self.percent.setStyleSheet("color: #FFFFFF; font-size: 30px; font-weight: 800; background: transparent;")
        progress_row.addWidget(self.progress_bar, 1)
        progress_row.addWidget(self.percent)
        root.addLayout(progress_row)

        self.status = QLabel("Подготовка...")
        self.status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status.setStyleSheet("color: #E8EEF7; font-size: 16px; font-weight: 700; background: transparent;")
        root.addWidget(self.status)

        self.rows: dict[str, LoadingStepRow] = {}
        rows = QVBoxLayout()
        rows.setContentsMargins(360, 8, 360, 0)
        rows.setSpacing(2)
        for step in STARTUP_STEPS:
            row = LoadingStepRow(step)
            self.rows[step.key] = row
            rows.addWidget(row)
        root.addLayout(rows)
        root.addStretch(1)

        footer = QHBoxLayout()
        footer.addStretch(1)
        footer_text = QLabel("WGA  |  Your system. Our care.")
        footer_text.setStyleSheet("color: #FAD37A; font-size: 18px; font-style: italic; background: transparent;")
        footer.addWidget(footer_text)
        root.addLayout(footer)

    def start(self) -> None:
        self._thread = QThread(self)
        self._worker = StartupWorker()
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.step_started.connect(self._step_started)
        self._worker.step_completed.connect(self._step_completed)
        self._worker.step_failed.connect(self._step_failed)
        self._worker.warning.connect(self._warnings.append)
        self._worker.finished.connect(self._worker_finished)
        self._worker.finished.connect(self._thread.quit)
        self._worker.finished.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.start()

    def complete_gui_step(self, key: str, detail: str, progress: int) -> None:
        self._step_started(key, self.rows[key].step.label)
        self._step_completed(key, detail, progress)

    def finish_success(self, window: object) -> None:
        self._step_started("ready", "Готово")
        self._step_completed("ready", "Началното меню е готово.", 100)
        self._pending_window = window
        self._maybe_finish_success()

    def _maybe_finish_success(self) -> None:
        if self._pending_window is None:
            return
        if self._current_progress < 100 or self._visible_timer.elapsed() < self._minimum_visible_ms:
            QTimer.singleShot(120, self._maybe_finish_success)
            return
        window = self._pending_window
        self._pending_window = None
        QTimer.singleShot(420, lambda: self._fade_out(window))

    def _step_started(self, key: str, label: str) -> None:
        row = self.rows.get(key)
        if row:
            row.set_status("active")
        self.status.setText(label)

    def _step_completed(self, key: str, detail: str, progress: int) -> None:
        row = self.rows.get(key)
        if row:
            row.set_status("done")
        self.status.setText(detail)
        self._target_progress = max(self._target_progress, progress)

    def _step_failed(self, key: str, label: str, detail: str, required: bool) -> None:
        row = self.rows.get(key)
        if row:
            row.set_status("error" if required else "warning")
        self.status.setText(f"{label}: {detail}")
        if required:
            self.failed.emit(detail)

    def _worker_finished(self, ok: bool) -> None:
        if ok:
            self.completed.emit(self._warnings)

    def _animate_progress(self) -> None:
        if self._current_progress == self._target_progress:
            return
        delta = max(1, int((self._target_progress - self._current_progress) * 0.22))
        self._current_progress = min(self._target_progress, self._current_progress + delta)
        self.progress_bar.set_value(self._current_progress)
        self.percent.setText(f"{self._current_progress}%")

    def _animate_waves(self) -> None:
        self._wave_shift = (self._wave_shift + 1) % 80
        self.update()

    def _fade_out(self, window: object) -> None:
        fade = QPropertyAnimation(self.opacity, b"opacity", self)
        fade.setDuration(420)
        fade.setStartValue(self.opacity.opacity())
        fade.setEndValue(0.0)
        fade.setEasingCurve(QEasingCurve.Type.InCubic)
        fade.finished.connect(lambda: self._finish_transition(window))
        fade.start(QPropertyAnimation.DeletionPolicy.DeleteWhenStopped)

    def _finish_transition(self, window: object) -> None:
        self.close()
        if hasattr(window, "show"):
            window.show()
        if hasattr(window, "raise_"):
            window.raise_()
        if hasattr(window, "activateWindow"):
            window.activateWindow()

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        rect = self.rect()
        painter.fillRect(rect, QColor("#03070D"))
        if not self._background.isNull():
            painter.drawPixmap(rect, self._background)
        if not self._planet.isNull():
            planet_width = rect.width() * 0.96
            planet_rect = QRectF((rect.width() - planet_width) / 2, rect.height() * -0.08, planet_width, rect.height() * 0.68)
            painter.drawPixmap(planet_rect.toRect(), self._planet)
        if not self._waves_1.isNull():
            y = int(rect.height() * 0.55)
            painter.setOpacity(0.82)
            painter.drawPixmap(QRectF(-90 + self._wave_shift, y, rect.width() + 180, rect.height() * 0.48).toRect(), self._waves_1)
        if not self._waves_2.isNull():
            y = int(rect.height() * 0.50)
            painter.setOpacity(0.72)
            painter.drawPixmap(QRectF(-100 - self._wave_shift, y, rect.width() + 200, rect.height() * 0.48).toRect(), self._waves_2)
        if not self._line.isNull():
            painter.setOpacity(0.95)
            painter.drawPixmap(QRectF(rect.width() * 0.78, rect.height() * 0.13, rect.width() * 0.16, 22).toRect(), self._line)
        painter.setOpacity(1.0)
        painter.fillRect(rect, QColor(0, 0, 0, 32))
        super().paintEvent(event)
