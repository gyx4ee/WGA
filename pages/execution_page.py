from __future__ import annotations

from PySide6.QtCore import QObject, Qt, QThread, QTimer, Signal
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import QCheckBox, QDialog, QFrame, QGridLayout, QHBoxLayout, QInputDialog, QLabel, QMessageBox, QProgressBar, QPushButton, QScrollArea, QVBoxLayout, QWidget

from services.paths import ui_asset
from services.action_runner import get_auto_installer_office_packages, get_auto_installer_packages, get_auto_installer_presets, run_action
from widgets.circular_progress import CircularProgress


def _steps_for_action(action_id: str, title: str) -> list[str]:
    text = f"{action_id} {title}".lower()
    if "office" in text:
        return [
            "Проверка за инсталиран Office пакет",
            "Търсене на OSPP.VBS / Click-to-Run компоненти",
            "Подготовка на Office командите",
            "Изпълнение на Office операцията",
            "Проверка на статуса след изпълнение",
        ]
    if "windows" in text or "activate" in text:
        return [
            "Проверка на Windows edition и build",
            "Проверка за администраторски права",
            "Подготовка на slmgr/cscript командите",
            "Изпълнение на activation workflow",
            "Проверка на лицензионния статус",
        ]
    if "driver" in text or "report" in text:
        return [
            "Събиране на системна информация",
            "Проверка на driver store",
            "Подготовка на output папките",
            "Изпълнение на driver/report операцията",
            "Финална проверка на резултата",
        ]
    if "language" in text or "keyboard" in text:
        return [
            "Проверка на текущите езикови настройки",
            "Зареждане на наличните keyboard layouts",
            "Подготовка на Windows language командите",
            "Прилагане на избраната настройка",
            "Проверка на активния layout",
        ]
    return [
        "Проверка на системното състояние",
        "Проверка на нужните компоненти",
        "Подготовка на командите",
        "Изпълнение на действието",
        "Финална проверка",
    ]


class ActionWorker(QObject):
    event = Signal(int, str, str, bool)
    done = Signal()

    def __init__(self, action_id: str, title: str, mode: str, selected_packages: list[str] | None = None) -> None:
        super().__init__()
        self.action_id = action_id
        self.title = title
        self.mode = mode
        self.selected_packages = selected_packages or []

    def run(self) -> None:
        for item in run_action(self.action_id, self.title, self.mode, self.selected_packages):
            self.event.emit(int(item["progress"]), str(item["status"]), str(item["detail"]), bool(item["finished"]))
        self.done.emit()


class ExecutionProgressPage(QScrollArea):
    close_requested = Signal()

    def __init__(
        self,
        title: str,
        description: str,
        action_id: str,
        mode: str = "action",
        parent=None,
        selected_installer_packages: list[str] | None = None,
        progress_only: bool = False,
    ) -> None:
        super().__init__(parent)
        self.title_text = title
        self.description_text = description
        self.action_id = action_id
        self.mode = mode
        self.progress_only = progress_only
        self.preselected_installer_packages = selected_installer_packages or []
        self.progress_value = 0
        self.thread: QThread | None = None
        self.worker: ActionWorker | None = None
        all_installer_packages = get_auto_installer_packages() if mode == "installer" else []
        if mode == "installer" and progress_only:
            selected_set = set(self.preselected_installer_packages)
            self.installer_packages = [(label, package_id) for label, package_id in all_installer_packages if package_id in selected_set]
        else:
            self.installer_packages = all_installer_packages
        self.installer_office_packages = get_auto_installer_office_packages() if mode == "installer" else []
        self.installer_presets = get_auto_installer_presets() if mode == "installer" else []
        self.steps = self._auto_installer_steps() if mode == "installer" else _steps_for_action(action_id, title)
        self.step_rows: list[tuple[QLabel, QProgressBar, QLabel]] = []
        self.installer_step_rows: dict[str, QFrame] = {}
        self.installer_checks: dict[str, QCheckBox] = {}
        self.installer_selection_status: QLabel | None = None

        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        host = QWidget()
        self.setWidget(host)
        layout = QVBoxLayout(host)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)

        hero = QFrame()
        hero.setObjectName("executionHero")
        hero_layout = QHBoxLayout(hero)
        hero_layout.setContentsMargins(26, 22, 26, 22)
        hero_layout.setSpacing(22)

        left = QVBoxLayout()
        brand = QHBoxLayout()
        logo = QLabel()
        logo.setPixmap(QPixmap(str(ui_asset("logo/wga_symbol.png"))).scaled(58, 58, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        brand.addWidget(logo)
        brand_text = QVBoxLayout()
        wga = QLabel("WGA")
        wga.setProperty("brandTitle", True)
        sub = QLabel("WinSys Guardian Advanced")
        sub.setProperty("muted", True)
        brand_text.addWidget(wga)
        brand_text.addWidget(sub)
        brand.addLayout(brand_text)
        brand.addStretch(1)
        left.addLayout(brand)

        title = QLabel("Инсталиране на приложения" if mode == "installer" else title)
        title.setProperty("executionTitle", True)
        title.setWordWrap(True)
        left.addWidget(title)
        desc = QLabel(self._subtitle())
        desc.setProperty("muted", True)
        desc.setWordWrap(True)
        left.addWidget(desc)

        self.overall = QProgressBar()
        self.overall.setRange(0, 100)
        self.overall.setValue(0)
        self.overall.setProperty("executionBar", True)
        left.addWidget(self.overall)

        metrics = QGridLayout()
        metrics.setSpacing(10)
        initial_total = str(len(self.installer_packages)) if mode == "installer" and progress_only else "0" if mode == "installer" else str(len(self.steps))
        self.metric_total = self._metric("Общо", initial_total)
        self.metric_done = self._metric("Завършени", "0")
        self.metric_waiting = self._metric("Оставащи", initial_total)
        self.metric_time = self._metric("Време", "00:00")
        for index, metric in enumerate((self.metric_total, self.metric_done, self.metric_waiting, self.metric_time)):
            metrics.addWidget(metric, 0, index)
        left.addLayout(metrics)
        hero_layout.addLayout(left, 1)

        self.circle = CircularProgress("WGA", accent="#20C5FF", size=210)
        hero_layout.addWidget(self.circle, 0, Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(hero)

        if mode == "installer" and not progress_only:
            selection_card = QFrame()
            selection_card.setProperty("card", True)
            selection_layout = QGridLayout(selection_card)
            selection_layout.setContentsMargins(16, 14, 16, 14)
            selection_layout.setSpacing(8)
            selection_title = QLabel("Избрани приложения")
            selection_title.setProperty("subtitle", True)
            selection_layout.addWidget(selection_title, 0, 0, 1, 3)
            preset_row = QHBoxLayout()
            for preset in self.installer_presets:
                button = QPushButton(str(preset["label"]))
                button.setProperty("secondary", True)
                button.clicked.connect(lambda _checked=False, current=preset: self._apply_installer_preset(current))
                preset_row.addWidget(button)
            clear_button = QPushButton("Размаркирай всички")
            clear_button.setProperty("secondary", True)
            clear_button.clicked.connect(self._clear_installer_selection)
            preset_row.addWidget(clear_button)
            preset_row.addStretch(1)
            selection_layout.addLayout(preset_row, 1, 0, 1, 3)
            self.installer_selection_status = QLabel("Нищо не е избрано. Избери пакет или маркирай отделни приложения.")
            self.installer_selection_status.setProperty("muted", True)
            self.installer_selection_status.setWordWrap(True)
            selection_layout.addWidget(self.installer_selection_status, 2, 0, 1, 3)
            for index, (label_text, package_id) in enumerate(self.installer_packages, start=1):
                checkbox = QCheckBox(label_text)
                checkbox.setChecked(False)
                checkbox.toggled.connect(self._update_installer_selection_status)
                self.installer_checks[package_id] = checkbox
                selection_layout.addWidget(checkbox, 3 + ((index - 1) // 3), (index - 1) % 3)
            layout.addWidget(selection_card)

        list_card = QFrame()
        list_card.setProperty("card", True)
        list_layout = QVBoxLayout(list_card)
        list_layout.setContentsMargins(16, 14, 16, 14)
        list_layout.setSpacing(8)
        list_title = QLabel("Детайли по изпълнението")
        list_title.setProperty("subtitle", True)
        list_layout.addWidget(list_title)
        if mode == "installer":
            self.empty_steps_label = QLabel("Избраните приложения ще се появят тук." if not progress_only else "Инсталацията започва...")
            self.empty_steps_label.setProperty("muted", True)
            list_layout.addWidget(self.empty_steps_label)
            for label_text, package_id in self.installer_packages:
                row = self._step_row(label_text)
                row.setVisible(progress_only)
                self.installer_step_rows[package_id] = row
                list_layout.addWidget(row)
            if progress_only:
                self.empty_steps_label.setVisible(False)
        else:
            self.empty_steps_label = None
            for step in self.steps:
                list_layout.addWidget(self._step_row(step))
        layout.addWidget(list_card)

        footer = QFrame()
        footer.setProperty("card", True)
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(16, 12, 16, 12)
        self.autostart = QCheckBox("Автоматично стартиране след успешна операция")
        self.autostart.setChecked(False)
        self.autostart.setVisible(mode != "installer" and not progress_only)
        footer_layout.addWidget(self.autostart)
        footer_layout.addStretch(1)
        self.start_button = QPushButton("Стартирай")
        self.start_button.clicked.connect(self.start)
        footer_layout.addWidget(self.start_button)
        self.cancel_button = QPushButton("Отмени")
        self.cancel_button.setProperty("secondary", True)
        self.cancel_button.clicked.connect(self.cancel)
        footer_layout.addWidget(self.cancel_button)
        layout.addWidget(footer)
        layout.addStretch(1)

    def _subtitle(self) -> str:
        if self.mode == "installer":
            return "Инсталират се избраните приложения с детайлен статус за всяка стъпка."
        return self.description_text or "Изпълнява се избраното WGA действие с детайлна информация."

    def _auto_installer_steps(self) -> list[str]:
        return [label for label, _package_id in self.installer_packages]

    def _apply_installer_preset(self, preset: dict[str, object]) -> None:
        package_ids = set(str(package_id) for package_id in preset.get("packages", []))
        for package_id, checkbox in self.installer_checks.items():
            checkbox.setChecked(package_id in package_ids)
        self._update_installer_selection_status()

    def _clear_installer_selection(self) -> None:
        for checkbox in self.installer_checks.values():
            checkbox.setChecked(False)
        self._update_installer_selection_status()

    def _update_installer_selection_status(self, *_args: object) -> None:
        if self.installer_selection_status is None:
            return
        selected = len(self._selected_installer_packages())
        total = len(self.installer_checks)
        self._sync_installer_step_rows()
        if selected:
            self.installer_selection_status.setText(f"Избрани са {selected} от {total} действия. Инсталацията ще започне само след натискане на Стартирай.")
        else:
            self.installer_selection_status.setText("Нищо не е избрано. Избери пакет или маркирай отделни приложения.")

    def _sync_installer_step_rows(self) -> None:
        if self.mode != "installer":
            return
        selected_ids = set(self._selected_installer_packages())
        for package_id, row in self.installer_step_rows.items():
            row.setVisible(package_id in selected_ids)
        if self.empty_steps_label is not None:
            self.empty_steps_label.setVisible(not selected_ids)
        selected_count = len(selected_ids)
        self.metric_total.value_label.setText(str(selected_count))  # type: ignore[attr-defined]
        self.metric_done.value_label.setText("0")  # type: ignore[attr-defined]
        self.metric_waiting.value_label.setText(str(selected_count))  # type: ignore[attr-defined]

    def _active_step_rows(self) -> list[tuple[QLabel, QProgressBar, QLabel]]:
        if self.mode != "installer":
            return self.step_rows
        selected_ids = self._selected_installer_packages()
        active_rows: list[tuple[QLabel, QProgressBar, QLabel]] = []
        for package_id in selected_ids:
            row = self.installer_step_rows.get(package_id)
            if row is None:
                continue
            for label, progress, status in self.step_rows:
                if label.parentWidget() is row:
                    active_rows.append((label, progress, status))
                    break
        return active_rows

    def _metric(self, title: str, value: str) -> QFrame:
        frame = QFrame()
        frame.setProperty("executionMetric", True)
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(12, 10, 12, 10)
        value_label = QLabel(value)
        value_label.setProperty("metricTitle", True)
        title_label = QLabel(title)
        title_label.setProperty("metricDetail", True)
        layout.addWidget(value_label)
        layout.addWidget(title_label)
        frame.value_label = value_label  # type: ignore[attr-defined]
        return frame

    def _step_row(self, title: str) -> QFrame:
        row = QFrame()
        row.setProperty("executionStep", True)
        layout = QHBoxLayout(row)
        layout.setContentsMargins(12, 8, 12, 8)
        label = QLabel(title)
        label.setMinimumWidth(250)
        progress = QProgressBar()
        progress.setRange(0, 100)
        progress.setValue(0)
        progress.setProperty("executionStepBar", True)
        status = QLabel("Изчаква...")
        status.setProperty("metricDetail", True)
        status.setMinimumWidth(120)
        layout.addWidget(label, 2)
        layout.addWidget(progress, 2)
        layout.addWidget(status, 1)
        self.step_rows.append((label, progress, status))
        return row

    def start(self) -> None:
        self.progress_value = 0
        if self.mode == "installer":
            selected_packages = self._selected_installer_packages()
            if not selected_packages:
                if self.installer_selection_status is not None:
                    self.installer_selection_status.setText("Първо избери приложения или preset пакет, после натисни Стартирай.")
                self._set_active_status("Няма избрани приложения")
                return
            if not self.progress_only:
                selected_packages = self._review_installer_selection(selected_packages)
                if selected_packages is None:
                    return
                self._open_installer_progress_dialog(selected_packages)
                return
        else:
            selected_packages = []
        self.start_button.setEnabled(False)
        self.cancel_button.setText("Отмени")
        self._sync_installer_step_rows()
        for _label, progress, status in self._active_step_rows():
            progress.setValue(0)
            status.setText("Изчаква...")
        self.thread = QThread(self)
        self.worker = ActionWorker(self.action_id, self.title_text, self.mode, selected_packages)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.event.connect(self._apply_runner_event)
        self.worker.done.connect(self.thread.quit)
        self.worker.done.connect(self.worker.deleteLater)
        self.thread.finished.connect(self._runner_finished)
        self.thread.start()

    def _selected_installer_packages(self) -> list[str]:
        if self.mode == "installer" and self.progress_only:
            return list(self.preselected_installer_packages)
        return [package_id for package_id, checkbox in self.installer_checks.items() if checkbox.isChecked()]

    def _open_installer_progress_dialog(self, selected_packages: list[str]) -> None:
        dialog = QDialog(self.window())
        dialog.setWindowTitle("Инсталиране на приложения")
        dialog.setModal(False)
        dialog.resize(980, 760)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(10, 10, 10, 10)
        page = ExecutionProgressPage(
            "Инсталиране на приложения",
            "Избраните приложения се инсталират едно по едно в silent режим.",
            self.action_id,
            self.mode,
            dialog,
            selected_installer_packages=selected_packages,
            progress_only=True,
        )
        layout.addWidget(page)
        page.close_requested.connect(dialog.close)
        dialog.finished.connect(dialog.deleteLater)
        dialog.show()
        frame = dialog.frameGeometry()
        screen = dialog.screen()
        if screen is not None:
            frame.moveCenter(screen.availableGeometry().center())
            dialog.move(frame.topLeft())
        dialog.raise_()
        dialog.activateWindow()
        QTimer.singleShot(250, page.start)

    def _review_installer_selection(self, selected_packages: list[str]) -> list[str] | None:
        selected = set(selected_packages)
        if self.installer_office_packages and not any(package_id in selected for _label, package_id in self.installer_office_packages):
            updated = self._prompt_for_missing_office(selected)
            if updated is None:
                return None
            selected = updated
        if "auto_windows_activation" not in selected:
            if not self._prompt_for_activation(
                "auto_windows_activation",
                "Windows активация не е избрана",
                "Windows активацията не е отметната. WGA може автоматично да провери Windows версията, build-а и записания ключ.\n\nДа добавя ли Windows активация към края на задачите?",
            ):
                return None
            selected = set(self._selected_installer_packages())
        if "auto_office_activation" not in selected:
            if not self._prompt_for_activation(
                "auto_office_activation",
                "Office активация не е избрана",
                "Office активацията не е отметната. WGA ще провери инсталирания Office пакет и записания ключ, а ако липсва ключ ще покаже съобщение.\n\nДа добавя ли Office активация към края на задачите?",
            ):
                return None
            selected = set(self._selected_installer_packages())
        self._update_installer_selection_status()
        return self._selected_installer_packages()

    def _prompt_for_missing_office(self, selected: set[str]) -> set[str] | None:
        labels = ["Продължи без Office"]
        label_to_package = {label: package_id for label, package_id in self.installer_office_packages}
        labels.extend(label_to_package)
        choice, ok = QInputDialog.getItem(
            self,
            "Office не е избран",
            "Не е избран Office пакет за инсталация. Ако Office вече е инсталиран, можеш да продължиш без него.\n\nИзбери Office версия или продължи без Office:",
            labels,
            0,
            False,
        )
        if not ok:
            return None
        package_id = label_to_package.get(choice)
        if package_id and package_id in self.installer_checks:
            self.installer_checks[package_id].setChecked(True)
            selected.add(package_id)
        return selected

    def _prompt_for_activation(self, package_id: str, title: str, message: str) -> bool:
        result = QMessageBox.question(
            self,
            title,
            message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Yes,
        )
        if result == QMessageBox.StandardButton.Cancel:
            return False
        if result == QMessageBox.StandardButton.Yes and package_id in self.installer_checks:
            self.installer_checks[package_id].setChecked(True)
        return True

    def cancel(self) -> None:
        if self.thread and self.thread.isRunning():
            self.cancel_button.setText("Спира се...")
            self.cancel_button.setEnabled(False)
            return
        if self.cancel_button.text() == "Затвори":
            self.close_requested.emit()
            return
        self.start_button.setEnabled(True)
        self.cancel_button.setText("Отмени")
        if self.progress_value < 100:
            self._set_active_status("Отменено")

    def _apply_runner_event(self, progress: int, status_text: str, detail: str, finished: bool) -> None:
        self.progress_value = max(0, min(100, progress))
        self.overall.setValue(self.progress_value)
        self.circle.set_value(self.progress_value)
        rows = self._active_step_rows()
        if not rows:
            return
        completed = int((self.progress_value / 100) * len(rows))
        active = min(completed, len(rows) - 1)
        for index, (_label, progress, status) in enumerate(rows):
            if index < completed:
                progress.setValue(100)
                status.setText("Завършено")
            elif index == active:
                local = int(((self.progress_value / 100) * len(rows) - index) * 100)
                progress.setValue(max(8, min(99, local)))
                status.setText(status_text)
            else:
                progress.setValue(0)
                status.setText("Изчаква...")
        self.metric_done.value_label.setText(str(min(completed, len(rows))))  # type: ignore[attr-defined]
        self.metric_waiting.value_label.setText(str(max(0, len(rows) - completed)))  # type: ignore[attr-defined]
        self.metric_time.value_label.setText(f"{self.progress_value}%")  # type: ignore[attr-defined]
        if rows:
            rows[active][2].setText(status_text)
            rows[active][0].setText(detail[:120] + ("..." if len(detail) > 120 else ""))
        if finished:
            self.progress_value = 100
            self.overall.setValue(100)
            self.circle.set_value(100)

    def _set_active_status(self, text: str) -> None:
        for _label, progress, status in self.step_rows:
            if 0 < progress.value() < 100:
                status.setText(text)
                return

    def _runner_finished(self) -> None:
        self.start_button.setEnabled(True)
        self.cancel_button.setEnabled(True)
        self.cancel_button.setText("Затвори")
        if self.thread:
            self.thread.deleteLater()
            self.thread = None
        self.worker = None
