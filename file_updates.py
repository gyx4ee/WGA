"""Manifest-based updates shared by the client and hosting publisher."""
from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path, PurePosixPath
from urllib.parse import quote, urljoin, urlsplit
from urllib.request import Request, urlopen


FORMAT = 'wga-files-v1'
PRESERVED = {'settings.json', '.wga_secure_store.json', '.wga_secure_store.json.bak'}
EXCLUDED = {'.git', '__pycache__', 'logs', 'Installers', 'build', 'dist', 'installer-output'}


def digest_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def validate_manifest(data: object) -> list[dict]:
    if not isinstance(data, dict) or data.get('format') != FORMAT:
        raise ValueError('Неподдържан файлов manifest за актуализация.')
    if not re.fullmatch(r'\d+(?:\.\d+){1,3}', str(data.get('version', ''))):
        raise ValueError('Невалидна версия в manifest.')
    entries = data.get('files')
    if not isinstance(entries, list) or not 1 <= len(entries) <= 20000:
        raise ValueError('Невалиден списък с файлове.')
    seen = set()
    for entry in entries:
        if not isinstance(entry, dict):
            raise ValueError('Невалиден запис за файл.')
        name = entry.get('path', '')
        if not isinstance(name, str) or not name or '\\' in name or re.search(r'[<>:"|?*\x00-\x1f]', name):
            raise ValueError('Невалиден файлов път.')
        parts = name.split('/')
        if any(part in {'', '.', '..'} or part.endswith((' ', '.')) for part in parts):
            raise ValueError('Небезопасен файлов път.')
        if any(re.fullmatch(r'(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?', part, re.I) for part in parts):
            raise ValueError('Запазено файлово име.')
        if name.casefold() in seen:
            raise ValueError('Дублиращ се файлов път.')
        seen.add(name.casefold())
        if parts[-1].casefold() in {p.casefold() for p in PRESERVED} or any(p.casefold() in {e.casefold() for e in EXCLUDED} for p in parts[:-1]):
            raise ValueError('Актуализацията не може да заменя потребителски данни.')
        if type(entry.get('size')) is not int or not 0 <= entry['size'] <= 4 * 1024**3:
            raise ValueError('Невалиден размер на файл.')
        if not re.fullmatch(r'[0-9a-fA-F]{64}', str(entry.get('sha256', ''))):
            raise ValueError('Липсва валидна SHA-256 сума.')
    if 'wga.exe' not in seen or '_internal/version.json' not in seen:
        raise ValueError('Липсват WGA.exe или информация за версията.')
    return entries


def build_manifest(root: Path, version: str) -> dict:
    files = []
    for path in sorted(root.rglob('*')):
        if path.is_symlink():
            raise ValueError(f'Символна връзка в release: {path.name}')
        relative = path.relative_to(root)
        if path.is_file() and not any(p in EXCLUDED for p in relative.parts) and path.name not in PRESERVED:
            files.append({'path': relative.as_posix(), 'size': path.stat().st_size, 'sha256': digest_file(path)})
    data = {'format': FORMAT, 'version': version, 'files': files}
    validate_manifest(data)
    return data


def _open(url: str):
    response = urlopen(Request(url, headers={'User-Agent': 'WGA-File-Updater/1.0', 'Cache-Control': 'no-cache'}), timeout=45)
    original, final = urlsplit(url), urlsplit(response.geturl())
    if (final.scheme, final.netloc) != (original.scheme, original.netloc):
        response.close()
        raise ValueError('Източникът пренасочва към друг сървър.')
    return response


def download_file_update(manifest_url: str, staging: Path, target: Path, progress=None) -> Path:
    parts = urlsplit(manifest_url)
    if parts.scheme != 'https' and not (parts.scheme == 'http' and parts.hostname in {'127.0.0.1', 'localhost', '::1'}):
        raise ValueError('Файловите актуализации изискват HTTPS.')
    if parts.username or parts.password or not parts.netloc or parts.query or parts.fragment:
        raise ValueError('Невалиден адрес за manifest.')
    with _open(manifest_url) as response:
        raw = response.read(8 * 1024 * 1024 + 1)
    if len(raw) > 8 * 1024 * 1024:
        raise ValueError('Manifest файлът е прекалено голям.')
    data = json.loads(raw.decode('utf-8-sig'))
    entries = validate_manifest(data)
    staging.mkdir(parents=True, exist_ok=True)
    total = max(1, sum(entry['size'] for entry in entries))
    completed = 0
    for entry in entries:
        name, size, expected = entry['path'], entry['size'], entry['sha256'].lower()
        local = target.joinpath(*PurePosixPath(name).parts)
        if not local.resolve().is_relative_to(target.resolve()):
            raise ValueError('Локалният файлов път излиза извън папката на WGA.')
        if not local.is_symlink() and local.is_file() and local.stat().st_size == size and digest_file(local) == expected:
            completed += size
            if progress:
                progress(min(90, 5 + int(completed * 85 / total)))
            continue
        output = staging.joinpath(*PurePosixPath(name).parts)
        output.parent.mkdir(parents=True, exist_ok=True)
        count = 0
        digest = hashlib.sha256()
        try:
            with _open(urljoin(manifest_url, 'files/' + quote(name, safe='/'))) as response, output.open('wb') as stream:
                while chunk := response.read(256 * 1024):
                    count += len(chunk)
                    if count > size:
                        raise ValueError(f'Неправилен размер: {name}')
                    digest.update(chunk)
                    stream.write(chunk)
                    if progress:
                        progress(min(90, 5 + int((completed + count) * 85 / total)))
            if count != size or digest.hexdigest() != expected:
                raise ValueError(f'Неуспешна проверка на файла: {name}')
        except Exception:
            output.unlink(missing_ok=True)
            raise
        completed += size
    version_path = staging / '_internal' / 'version.json'
    if not version_path.exists():
        version_path = target / '_internal' / 'version.json'
    if json.loads(version_path.read_text(encoding='utf-8-sig')).get('version') != data['version']:
        raise ValueError('Версията на файловете не съвпада с manifest.')
    return staging
