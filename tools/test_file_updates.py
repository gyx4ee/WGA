from __future__ import annotations

import copy
import json
import sys
import tempfile
import threading
import unittest
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from file_updates import build_manifest, download_file_update, validate_manifest
from self_updater import prepare_update_install
from tools.prepare_wsga_release import BASE_URL, publish_directory
from update_checker import check_for_updates


class FileUpdateTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        root = Path(self.temp.name)
        self.build = root / 'build'
        (self.build / '_internal').mkdir(parents=True)
        self.metadata = {'version': '0.2.69', 'version_info_url': BASE_URL + 'version.json',
            'notes': 'File update test', 'changelog': []}
        (self.build / 'WGA.exe').write_bytes(b'new executable')
        (self.build / '_internal' / 'version.json').write_text(json.dumps(self.metadata), encoding='utf-8')
        (self.build / '_internal' / 'unchanged.dll').write_bytes(b'unchanged')
        self.web = root / 'web'
        publish_directory(self.build, self.web, self.metadata)
        self.requests = []
        requests = self.requests

        class Handler(SimpleHTTPRequestHandler):
            def do_GET(self):
                requests.append(self.path)
                return super().do_GET()

            def log_message(self, *args):
                pass

        server = ThreadingHTTPServer(('127.0.0.1', 0), partial(Handler, directory=str(self.web)))
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(lambda: (server.shutdown(), thread.join(), server.server_close()))
        self.base = f'http://127.0.0.1:{server.server_port}/'
        self.url = self.base + 'releases/0.2.69/manifest.json'
        self.target = root / 'installed'
        (self.target / '_internal').mkdir(parents=True)
        (self.target / '_internal' / 'unchanged.dll').write_bytes(b'unchanged')
        (self.target / 'WGA.exe').write_bytes(b'old executable')
        (self.target / 'settings.json').write_text('private settings')
        self.staging = root / 'staged'

    def test_downloads_only_changed_files_and_never_mutates_installation(self):
        progress = []
        download_file_update(self.url, self.staging, self.target, progress.append)
        self.assertEqual((self.staging / 'WGA.exe').read_bytes(), b'new executable')
        self.assertFalse(any('unchanged.dll' in path for path in self.requests))
        self.assertEqual((self.target / 'WGA.exe').read_bytes(), b'old executable')
        self.assertEqual((self.target / 'settings.json').read_text(), 'private settings')
        self.assertEqual(progress, sorted(progress))

    def test_corruption_is_rejected_before_replacement(self):
        (self.web / 'releases/0.2.69/files/WGA.exe').write_bytes(b'bad executable')
        with self.assertRaisesRegex(ValueError, 'проверка'):
            download_file_update(self.url, self.staging, self.target)
        self.assertEqual((self.target / 'WGA.exe').read_bytes(), b'old executable')
        self.assertFalse((self.staging / 'WGA.exe').exists())

    def test_manifest_rejects_traversal_duplicates_and_user_data(self):
        manifest = build_manifest(self.build, '0.2.69')
        for path in ('../settings.json', '/outside', 'C:/evil', '_internal/../../evil', 'SETTINGS.json', 'CON', 'a\\b'):
            bad = copy.deepcopy(manifest)
            bad['files'][0]['path'] = path
            with self.subTest(path=path), self.assertRaises(ValueError):
                validate_manifest(bad)
        bad = copy.deepcopy(manifest)
        bad['files'].append(dict(bad['files'][0]))
        with self.assertRaises(ValueError):
            validate_manifest(bad)

    def test_missing_file_rejects_the_update(self):
        (self.web / 'releases/0.2.69/files/WGA.exe').unlink()
        with self.assertRaises(Exception):
            download_file_update(self.url, self.staging, self.target)
        self.assertEqual((self.target / 'WGA.exe').read_bytes(), b'old executable')

    def test_metadata_preserves_manifest_url_for_existing_ui(self):
        result = check_for_updates('0.2.68', self.base + 'version.json')
        self.assertEqual(result.status, 'update_available')
        self.assertTrue(result.package_url.endswith('/manifest.json'))
        self.assertTrue(result.package_url.startswith(BASE_URL))

    def test_prepare_creates_helper_without_downloading_a_zip(self):
        work = Path(self.temp.name) / 'work'
        work.mkdir()
        with patch('self_updater.tempfile.mkdtemp', return_value=str(work)):
            helper = prepare_update_install(package_url=self.url, target_root=self.target,
                restart_command=[str(self.target / 'WGA.exe')])
        self.assertTrue(helper.exists())
        self.assertFalse(list(work.rglob('*.zip')))
        self.assertEqual((self.target / 'WGA.exe').read_bytes(), b'old executable')

    def test_no_plain_http_for_remote_updates(self):
        with self.assertRaisesRegex(ValueError, 'HTTPS'):
            download_file_update('http://wsga.eu/updates/manifest.json', self.staging, self.target)


if __name__ == '__main__':
    unittest.main()
