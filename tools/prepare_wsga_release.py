"""Build and prepare an uncompressed, manually uploaded WSGA release."""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from file_updates import build_manifest


BASE_URL = 'https://wsga.eu/updates/wga/'
HTACCESS = '''Options -Indexes
<FilesMatch "^\\.">
    Require all denied
</FilesMatch>
<IfModule mod_headers.c>
    <Files "version.json">
        Header set Cache-Control "no-cache, no-store, must-revalidate"
    </Files>
</IfModule>
'''


def write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')


def publish_directory(build_dir: Path, output: Path, metadata: dict) -> Path:
    version = metadata['version']
    embedded = json.loads((build_dir / '_internal' / 'version.json').read_text(encoding='utf-8-sig'))
    if embedded.get('version') != version or embedded.get('version_info_url') != BASE_URL + 'version.json':
        raise ValueError('Билдът не съвпада с версията или използва стария update адрес.')
    # Older installations read this root metadata before the bundled copy.
    shutil.copy2(build_dir / '_internal' / 'version.json', build_dir / 'version.json')
    manifest = build_manifest(build_dir, version)
    release = output / 'releases' / version
    if release.exists():
        raise ValueError(f'Папката за тази версия вече съществува: {release}')
    files_root = release / 'files'
    for entry in manifest['files']:
        destination = files_root / entry['path']
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(build_dir / entry['path'], destination)
    write_json(release / 'manifest.json', manifest)
    metadata = dict(metadata)
    metadata.update(version_info_url=BASE_URL + 'version.json', download_url='https://wsga.eu/',
        package_url=BASE_URL + f'releases/{version}/manifest.json', update_format='wga-files-v1')
    write_json(output / 'version.json', metadata)
    (output / '.htaccess').write_text(HTACCESS, encoding='utf-8')
    print(f'Готово: {output}\nФайлове: {len(manifest["files"])}\nРазмер: {sum(e["size"] for e in manifest["files"]) / 1048576:.1f} MB')
    return release


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--build', action='store_true', help='Build WGA and agent before preparing files')
    parser.add_argument('--version', help='New version, e.g. 0.2.70 (required with --build)')
    parser.add_argument('--notes', help='Release notes for this version')
    parser.add_argument('--build-dir', type=Path)
    parser.add_argument('--output', type=Path, default=ROOT / 'hosting-upload' / 'updates' / 'wga')
    args = parser.parse_args()
    metadata_path = ROOT / 'version.json'
    metadata = json.loads(metadata_path.read_text(encoding='utf-8-sig'))
    if args.build:
        import re
        if not args.version or not re.fullmatch(r'\d+(?:\.\d+){1,3}', args.version):
            parser.error('--build изисква валидна --version')
        metadata.update(version=args.version, version_info_url=BASE_URL + 'version.json',
            package_url=BASE_URL + f'releases/{args.version}/manifest.json',
            download_url='https://wsga.eu/', update_format='wga-files-v1')
        if args.notes:
            metadata['notes'] = args.notes
            entry = f'v{args.version} - {args.notes}'
            metadata['changelog'] = [entry] + [c for c in metadata.get('changelog', []) if not c.startswith(f'v{args.version} - ')]
        write_json(metadata_path, metadata)
        dist = ROOT / 'dist' / args.version
        for spec in ('WGA.spec', 'wga_agent.spec'):
            subprocess.run([sys.executable, '-m', 'PyInstaller', '--noconfirm',
                '--distpath', str(dist), '--workpath', str(ROOT / 'build' / args.version), spec], cwd=ROOT, check=True)
        shutil.copy2(dist / 'wga_agent.exe', dist / 'WGA' / 'wga_agent.exe')
        args.build_dir = dist / 'WGA'
    build_dir = args.build_dir or ROOT / 'dist' / metadata['version'] / 'WGA'
    publish_directory(build_dir, args.output, metadata)


if __name__ == '__main__':
    main()
