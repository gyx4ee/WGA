from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication
from pages.software_page import PackageWorker, SoftwarePage, local_action
from resource_manager import ResourceItem


class SoftwarePageTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.item = ResourceItem('office_2021_offline', 'Office 2021', 'Office',
            ('Office2021/setup.exe', 'Office2021/Configuration.xml'),
            'https://example.com/office.zip', 'office.zip', True)
        self.check = SimpleNamespace(item=self.item, available=False, missing_files=())
        self.status = SimpleNamespace(checks=[self.check], installers_root=Path('Installers'),
            total=1, available=0, missing=1, downloadable_missing=1)

    def test_office_uses_configured_installer(self):
        self.assertEqual(local_action(self.item), 'install_office_2021_offline')

    def test_catalog_starts_unchecked_and_filters(self):
        with patch('pages.software_page.check_resource_status', return_value=self.status):
            page = SoftwarePage()
            self.assertEqual(page.table.item(0, 0).checkState(), Qt.CheckState.Unchecked)
            page.search.setText('no matching package')
            self.assertTrue(page.table.isRowHidden(0))
            page.deleteLater()

    def test_download_reports_progress_before_completion(self):
        worker = PackageWorker('download', [self.item])
        updates, results = [], []
        worker.update.connect(lambda *event: updates.append(event))
        worker.result.connect(results.append)

        def transfer(root, item, callback):
            callback(50, 100, item.name, 10, 5)
            self.assertTrue(any('50%' in title for _, title, _ in updates))
            self.check.available = True

        with patch('pages.software_page.download_resource', side_effect=transfer), patch(
            'pages.software_page.check_resource_status', return_value=self.status):
            worker.run()
        self.assertEqual(results, ['Операцията завърши.'])

    def test_install_failure_is_not_reported_as_success(self):
        worker = PackageWorker('install', [('Test app', 'Test.App')])
        results = []
        worker.result.connect(results.append)
        with patch('pages.software_page.run_action', return_value=iter([
            {'progress': 100, 'status': 'Има грешки', 'detail': 'failed', 'finished': False}])):
            worker.run()
        self.assertIn('Завършено с грешки', results[0])


if __name__ == '__main__':
    unittest.main()
