from __future__ import annotations

import re
import urllib.request

from PySide6.QtCore import QThread, Qt, Signal
from PySide6.QtGui import QBrush
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QDialog, QHBoxLayout,
    QGridLayout, QHeaderView, QLabel, QLineEdit, QPlainTextEdit, QProgressBar, QPushButton,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget)

from office_installers import OFFICE_OFFLINE_INSTALLERS
from resource_manager import check_resource_status, download_resource
from services.action_runner import AUTO_INSTALLER_PACKAGES, LOCAL_INSTALLER_TASKS, run_action
from services.paths import project_root
from ui.dialogs import modern_confirm, modern_message


def local_action(item):
    for installer in OFFICE_OFFLINE_INSTALLERS.values():
        if f"{installer.folder}/setup.exe" in item.required_files:
            return installer.action_id
    for action_id, spec in LOCAL_INSTALLER_TASKS.items():
        if spec['resource_id'] == item.resource_id:
            return action_id
    return f"resource_{item.resource_id}"


class PackageWorker(QThread):
    update = Signal(int, str, str)
    result = Signal(str)

    def __init__(self, mode, items, parent=None):
        super().__init__(parent)
        self.mode, self.items = mode, items

    def run(self):
        try:
            errors = []
            for index, item in enumerate(self.items):
                count = len(self.items)
                self.update.emit(int(index * 100 / count), f"Пакет {index + 1}/{count}", '')
                try:
                    if self.mode == 'online_check':
                        request = urllib.request.Request(item.url, method='HEAD', headers={'User-Agent': 'WGA-Resource-Downloader/1.0'})
                        with urllib.request.urlopen(request, timeout=15) as response:
                            size = int(response.headers.get('Content-Length') or 0)
                            self.update.emit(int((index + 1) * 100 / count), item.name,
                                f"Източник: Наличен | {size / 1048576:.1f} MB | HTTP {response.status}")
                    elif self.mode == 'download':
                        self.update.emit(int(index * 100 / count), item.name,
                            f"Източник: Наличен\nSHA-256: {'ще се провери' if item.sha256 else 'не е зададен в каталога'}")

                        def progress(done, total, name, speed=0.0, eta=0, phase=''):
                            percent = min(100, int(done * 100 / total)) if total else 0
                            self.update.emit(int((index + percent / 100) * 100 / count), f"{name} | {phase or 'Изтегляне'} | {percent if total else '?'}%",
                                f"{done / 1048576:.1f} / {f'{total / 1048576:.1f}' if total else '?'} MB | "
                                f"{speed / 1048576:.2f} MB/s | Остава: {eta // 60}:{eta % 60:02d}")

                        download_resource(project_root(), item, progress)
                        check = next(c for c in check_resource_status(project_root()).checks if c.item.resource_id == item.resource_id)
                        if not check.available:
                            raise RuntimeError('След изтеглянето липсват: ' + ', '.join(map(str, check.missing_files)))
                        self.update.emit(int((index + 1) * 100 / count), item.name, 'Пакетът е готов за локална инсталация.')
                    else:
                        label, action = item
                        for event in run_action('auto_installer', label, 'installer', [action]):
                            self.update.emit(int(event['progress']), str(event['status']), str(event['detail']))
                            if int(event['progress']) == 100 and not event.get('finished'):
                                errors.append(label)
                except Exception as exc:
                    errors.append(f"{getattr(item, 'name', str(item))}: {exc}")
                    self.update.emit(int((index + 1) * 100 / count), 'Грешка', errors[-1])
            self.result.emit('Завършено с грешки:\n' + '\n'.join(errors) if errors else 'Операцията завърши.')
        except Exception as exc:
            self.result.emit(f'Грешка: {exc}')


class PackageProgressDialog(QDialog):
    def __init__(self, mode, items, parent):
        super().__init__(parent)
        self.setWindowTitle('WGA | Софтуер и пакети')
        self.resize(800, 520)
        layout = QVBoxLayout(self)
        self.title = QLabel('Подготовка')
        self.title.setWordWrap(True)
        self.title.setProperty('subtitle', True)
        layout.addWidget(self.title)
        self.bar = QProgressBar()
        layout.addWidget(self.bar)
        self.detail = QLabel()
        self.detail.setWordWrap(True)
        layout.addWidget(self.detail)
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(1500)
        layout.addWidget(self.log, 1)
        self.close_button = QPushButton('Затвори')
        self.close_button.setEnabled(False)
        self.close_button.clicked.connect(self.close)
        layout.addWidget(self.close_button)
        self.worker = PackageWorker(mode, items, self)
        self.worker.update.connect(self.on_update)
        self.worker.result.connect(self.on_result)
        self.worker.finished.connect(lambda: self.close_button.setEnabled(True))

    def on_update(self, percent, title, detail):
        detail = re.sub(r'https?://\S+', 'онлайн източник', detail)
        self.bar.setValue(percent)
        self.title.setText(title)
        self.detail.setText(detail)
        if detail:
            self.log.appendPlainText(f'{title}\n{detail}')

    def on_result(self, message):
        message = re.sub(r'https?://\S+', 'онлайн източник', message)
        self.title.setText(message)
        self.log.appendPlainText(message)
        self.parent().refresh()

    def closeEvent(self, event):
        if self.worker.isRunning():
            event.ignore()
        else:
            event.accept()


class SoftwarePage(QWidget):
    navigate = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        toolbar = QGridLayout()
        self.mode = QComboBox()
        self.mode.addItems(['Локални пакети / офлайн', 'Онлайн приложения / WinGet'])
        self.mode.currentIndexChanged.connect(self.refresh)
        toolbar.addWidget(self.mode, 0, 0, 1, 2)
        self.search = QLineEdit()
        self.search.setPlaceholderText('Търсене по име')
        self.search.textChanged.connect(self.filter_rows)
        toolbar.addWidget(self.search, 1, 0)
        refresh = QPushButton('Провери локално')
        refresh.clicked.connect(self.refresh)
        toolbar.addWidget(refresh, 1, 1)
        layout.addLayout(toolbar)
        self.summary = QLabel()
        self.summary.setWordWrap(True)
        layout.addWidget(self.summary)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(['Пакет', 'Локално', 'Размер', 'Онлайн източник'])
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.cellClicked.connect(self.toggle_row)
        self.table.verticalHeader().setDefaultSectionSize(38)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table, 1)
        actions = QGridLayout()
        self.check_button = QPushButton('Провери онлайн')
        self.check_button.clicked.connect(lambda: self.start('online_check'))
        self.download_button = QPushButton('Изтегли избраните')
        self.download_button.clicked.connect(lambda: self.start('download'))
        install = QPushButton('Инсталирай избраните')
        install.clicked.connect(lambda: self.start('install'))
        clear = QPushButton('Размаркирай всички')
        clear.clicked.connect(self.clear_selection)
        for index, button in enumerate((self.check_button, self.download_button, install, clear)):
            actions.addWidget(button, index // 2, index % 2)
        layout.addLayout(actions)
        self.refresh()

    def refresh(self):
        self.table.setRowCount(0)
        self.checkboxes = []
        self.entries = []
        offline = self.mode.currentIndex() == 0
        self.check_button.setEnabled(offline)
        self.download_button.setEnabled(offline)
        status = check_resource_status(project_root())
        self.summary.setText(f'Installers: {status.installers_root}\nНалични: {status.available}/{status.total} | Липсват: {status.missing} | С адрес за теглене: {status.downloadable_missing}')
        if offline:
            self.entries = sorted(status.checks, key=lambda c: c.item.name.casefold())
            values = [(c.item.name, 'Наличен' if c.available else 'Липсва',
                f'{c.item.size_bytes / 1048576:.1f} MB' if c.item.size_bytes else 'Неизвестен',
                'Наличен' if c.item.url else 'Няма източник') for c in self.entries]
        else:
            self.entries = sorted(AUTO_INSTALLER_PACKAGES, key=lambda x: x[0].casefold())
            values = [(name, 'Онлайн', 'От WinGet', 'Наличен') for name, package in self.entries]
        for row, values_row in enumerate(values):
            self.table.insertRow(row)
            for col, value in enumerate(values_row):
                cell = QTableWidgetItem(value)
                if col == 0:
                    cell.setForeground(QBrush(Qt.GlobalColor.transparent))
                    if offline:
                        cell.setToolTip('\n'.join(self.entries[row].item.required_files))
                self.table.setItem(row, col, cell)
            checkbox = QCheckBox(values_row[0])
            checkbox.setCursor(Qt.CursorShape.PointingHandCursor)
            checkbox.setToolTip(self.table.item(row, 0).toolTip())
            self.checkboxes.append(checkbox)
            self.table.setCellWidget(row, 0, checkbox)
        self.filter_rows()

    def toggle_row(self, row, column):
        checkbox = self.checkboxes[row]
        checkbox.setChecked(not checkbox.isChecked())

    def filter_rows(self):
        query = self.search.text().casefold()
        for row in range(self.table.rowCount()):
            self.table.setRowHidden(row, query not in self.table.item(row, 0).text().casefold())

    def clear_selection(self):
        for checkbox in self.checkboxes:
            checkbox.setChecked(False)

    def start(self, mode):
        selected = [entry for row, entry in enumerate(self.entries)
            if self.checkboxes[row].isChecked()]
        if not selected:
            modern_message(self, 'Няма избор', 'Избери поне един пакет.')
            return
        offline = self.mode.currentIndex() == 0
        if mode == 'install' and offline:
            if any(not entry.available for entry in selected):
                modern_message(self, 'Липсващи файлове', 'Първо изтегли липсващите пакети и провери наличността отново.')
                return
            selected = [(entry.item.name, local_action(entry.item)) for entry in selected]
            if sum(action in OFFICE_OFFLINE_INSTALLERS for _, action in selected) > 1:
                modern_message(self, 'Office пакет', 'Избери само една версия на Office за инсталиране.')
                return
        elif mode in {'download', 'online_check'}:
            selected = [entry.item for entry in selected if entry.item.url and (mode != 'download' or not entry.available)]
        if not selected:
            modern_message(self, 'Няма пакети', 'Няма липсващи пакети с адрес за изтегляне в избора.')
            return
        if not modern_confirm(self, 'Софтуер и пакети', f'Да стартирам операцията за {len(selected)} пакета?'):
            return
        dialog = PackageProgressDialog(mode, selected, self)
        dialog.setModal(True)
        dialog.show()
        screen = self.screen()
        if screen:
            frame = dialog.frameGeometry()
            frame.moveCenter(screen.availableGeometry().center())
            dialog.move(frame.topLeft())
        dialog.worker.start()
        dialog.exec()
