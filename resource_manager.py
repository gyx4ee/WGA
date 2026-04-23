from __future__ import annotations

import json
import hashlib
import time
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path

from path_utils import resolve_installers_root


MANIFEST_FILE_NAME = "installers_manifest.json"


@dataclass(frozen=True)
class ResourceItem:
    resource_id: str
    name: str
    category: str
    required_files: tuple[str, ...]
    url: str
    target: str
    extract: bool
    size_bytes: int = 0
    sha256: str = ""


@dataclass(frozen=True)
class ResourceCheck:
    item: ResourceItem
    available: bool
    missing_files: tuple[Path, ...]


@dataclass(frozen=True)
class ResourceStatus:
    installers_root: Path
    total: int
    available: int
    missing: int
    downloadable_missing: int
    checks: tuple[ResourceCheck, ...]

    @property
    def complete(self) -> bool:
        return self.total > 0 and self.missing == 0

    @property
    def configured(self) -> bool:
        return self.total > 0


def manifest_path(program_root: Path) -> Path:
    portable_manifest = program_root / MANIFEST_FILE_NAME
    if portable_manifest.exists():
        return portable_manifest
    bundled_manifest = program_root / "_internal" / MANIFEST_FILE_NAME
    if bundled_manifest.exists():
        return bundled_manifest
    return portable_manifest


def load_resource_manifest(program_root: Path) -> list[ResourceItem]:
    path = manifest_path(program_root)
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []

    resources = data.get("resources", []) if isinstance(data, dict) else []
    items: list[ResourceItem] = []
    for raw_item in resources:
        if not isinstance(raw_item, dict):
            continue
        resource_id = str(raw_item.get("id", "")).strip()
        name = str(raw_item.get("name", resource_id)).strip()
        required_files = tuple(str(value).strip() for value in raw_item.get("required_files", []) if str(value).strip())
        if not resource_id or not required_files:
            continue
        items.append(
            ResourceItem(
                resource_id=resource_id,
                name=name or resource_id,
                category=str(raw_item.get("category", "Installers")).strip() or "Installers",
                required_files=required_files,
                url=str(raw_item.get("url", "")).strip(),
                target=str(raw_item.get("target", "")).strip(),
                extract=bool(raw_item.get("extract", False)),
                size_bytes=int(raw_item.get("size_bytes") or 0),
                sha256=str(raw_item.get("sha256", "")).strip(),
            )
        )
    return items


def check_resource_status(program_root: Path) -> ResourceStatus:
    installers_root = resolve_installers_root(program_root)
    checks: list[ResourceCheck] = []
    for item in load_resource_manifest(program_root):
        missing_files = tuple(
            installers_root / relative_path
            for relative_path in item.required_files
            if not (installers_root / relative_path).exists()
        )
        checks.append(ResourceCheck(item=item, available=not missing_files, missing_files=missing_files))

    missing_checks = [check for check in checks if not check.available]
    downloadable_missing = sum(1 for check in missing_checks if check.item.url)
    return ResourceStatus(
        installers_root=installers_root,
        total=len(checks),
        available=sum(1 for check in checks if check.available),
        missing=len(missing_checks),
        downloadable_missing=downloadable_missing,
        checks=tuple(checks),
    )


def download_resource(
    program_root: Path,
    item: ResourceItem,
    progress_callback: object | None = None,
) -> Path:
    if not item.url:
        raise ValueError(f"No download URL configured for {item.name}.")

    installers_root = resolve_installers_root(program_root)
    target_relative = item.target or Path(item.url).name or f"{item.resource_id}.download"
    target_path = installers_root / target_relative
    target_path.parent.mkdir(parents=True, exist_ok=True)

    start_time = time.monotonic()
    downloaded = 0
    request = urllib.request.Request(item.url, headers={"User-Agent": "WGA-Resource-Downloader/1.0"})
    with urllib.request.urlopen(request, timeout=60) as response:
        total_size = int(response.headers.get("Content-Length") or item.size_bytes or 0)
        with target_path.open("wb") as output:
            while True:
                chunk = response.read(1024 * 512)
                if not chunk:
                    break
                output.write(chunk)
                downloaded += len(chunk)
                if callable(progress_callback):
                    elapsed = max(0.001, time.monotonic() - start_time)
                    speed = downloaded / elapsed
                    eta = int((total_size - downloaded) / speed) if total_size and speed > 0 else 0
                    progress_callback(downloaded, total_size, item.name, speed, eta)

    if item.sha256:
        digest = hashlib.sha256()
        with target_path.open("rb") as downloaded_file:
            for chunk in iter(lambda: downloaded_file.read(1024 * 1024), b""):
                digest.update(chunk)
        if digest.hexdigest().upper() != item.sha256.upper():
            target_path.unlink(missing_ok=True)
            raise RuntimeError(f"SHA256 проверката е неуспешна за {item.name}.")

    if item.extract and zipfile.is_zipfile(target_path):
        extract_dir = target_path.parent
        with zipfile.ZipFile(target_path) as archive:
            archive.extractall(extract_dir)

    return target_path


def missing_resource_report(status: ResourceStatus) -> str:
    if not status.configured:
        return "Няма конфигуриран manifest за инсталационните ресурси."
    if status.complete:
        return "Всички описани инсталационни ресурси са налични."

    lines = [
        f"Налични: {status.available}/{status.total}",
        f"Липсват: {status.missing}",
        f"С адрес за изтегляне: {status.downloadable_missing}",
        "",
        "Липсващи ресурси:",
    ]
    for check in status.checks:
        if check.available:
            continue
        lines.append(f"- {check.item.name}")
        for missing_file in check.missing_files[:3]:
            lines.append(f"  {missing_file}")
        if len(check.missing_files) > 3:
            lines.append(f"  ... още {len(check.missing_files) - 3} файла")
    return "\n".join(lines)
