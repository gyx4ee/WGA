from __future__ import annotations

import os
import shutil
import tempfile
import zipfile
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlsplit, urlunsplit
from urllib.request import urlopen


PRESERVED_NAMES = {
    "settings.json",
    ".wga_secure_store.json",
    ".wga_secure_store.json.bak",
}

SKIPPED_DIRS = {
    ".git",
    "__pycache__",
    "build",
    "dist",
    "installer-output",
}


def _prepare_url(url: str) -> str:
    stripped_url = url.strip()
    if not stripped_url:
        raise ValueError("Липсва адрес за update пакет.")

    parts = urlsplit(stripped_url)
    if parts.scheme not in {"http", "https"} or not parts.netloc:
        raise ValueError("Адресът за update пакет трябва да бъде пълен http/https URL.")

    encoded_path = quote(parts.path, safe="/-._~")
    encoded_query = quote(parts.query, safe="=&-._~")
    encoded_fragment = quote(parts.fragment, safe="-._~")
    return urlunsplit((parts.scheme, parts.netloc, encoded_path, encoded_query, encoded_fragment))


def download_update_package(url: str, destination: Path, progress_callback=None) -> Path:
    destination.mkdir(parents=True, exist_ok=True)
    package_path = destination / "wga-update.zip"
    prepared_url = _prepare_url(url)

    try:
        with urlopen(prepared_url, timeout=30) as response:
            total = int(response.headers.get("Content-Length") or 0)
            downloaded = 0
            with package_path.open("wb") as output:
                while True:
                    chunk = response.read(1024 * 256)
                    if not chunk:
                        break
                    output.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback and total:
                        progress_callback(min(45, int(downloaded * 45 / total)))
    except HTTPError as exc:
        raise RuntimeError(f"Update пакетът не е достъпен. HTTP код: {exc.code}. URL: {prepared_url}") from exc
    except URLError as exc:
        raise RuntimeError(f"Update пакетът не може да бъде изтеглен: {exc.reason}") from exc

    if package_path.stat().st_size <= 0:
        raise RuntimeError("Update пакетът беше изтеглен като празен файл.")

    return package_path


def extract_update_package(package_path: Path, destination: Path, progress_callback=None) -> Path:
    extract_dir = destination / "extracted"
    if extract_dir.exists():
        shutil.rmtree(extract_dir)
    extract_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(package_path, "r") as archive:
        archive.extractall(extract_dir)

    entries = [item for item in extract_dir.iterdir() if item.name not in {"__MACOSX"}]
    source_root = entries[0] if len(entries) == 1 and entries[0].is_dir() else extract_dir

    if progress_callback:
        progress_callback(70)
    return source_root


def _cmd_quote(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'


def _restart_line(restart_command: list[str]) -> str:
    if not restart_command:
        return ""
    executable = _cmd_quote(restart_command[0])
    args = " ".join(_cmd_quote(item) for item in restart_command[1:])
    return f'start "" {executable} {args}'.rstrip()


def create_update_helper(
    *,
    source_root: Path,
    target_root: Path,
    restart_command: list[str],
    work_dir: Path,
) -> Path:
    helper_path = work_dir / "install_wga_update.cmd"
    restart_line = _restart_line(restart_command)
    skipped_dirs = " ".join(_cmd_quote(item) for item in sorted(SKIPPED_DIRS))
    preserved_files = " ".join(_cmd_quote(item) for item in sorted(PRESERVED_NAMES))
    log_path = target_root / "WGA-update.log"

    helper_text = f"""@echo off
chcp 65001 >nul
title WGA Update Installer
set "LOG_FILE={str(log_path)}"
echo Installing WinSys Guardian Advanced update... > "%LOG_FILE%"
echo Source: {_cmd_quote(str(source_root))} >> "%LOG_FILE%"
echo Target: {_cmd_quote(str(target_root))} >> "%LOG_FILE%"
echo Waiting for WGA to close...
timeout /t 4 /nobreak >nul
robocopy {_cmd_quote(str(source_root))} {_cmd_quote(str(target_root))} /E /R:5 /W:2 /XD {skipped_dirs} /XF {preserved_files} /LOG+:"%LOG_FILE%"
set "ROBOCOPY_CODE=%ERRORLEVEL%"
if %ROBOCOPY_CODE% GEQ 8 (
    echo Update failed. Robocopy code: %ROBOCOPY_CODE%
    echo Update failed. Robocopy code: %ROBOCOPY_CODE% >> "%LOG_FILE%"
    pause
    exit /b %ROBOCOPY_CODE%
)
echo Update installed.
echo Update installed. Robocopy code: %ROBOCOPY_CODE% >> "%LOG_FILE%"
{restart_line}
timeout /t 1 /nobreak >nul
exit /b 0
"""
    helper_path.write_text(helper_text, encoding="utf-8")
    return helper_path


def prepare_update_install(
    *,
    package_url: str,
    target_root: Path,
    restart_command: list[str],
    progress_callback=None,
) -> Path:
    work_dir = Path(tempfile.mkdtemp(prefix="wga_update_"))
    if progress_callback:
        progress_callback(5)

    package_path = download_update_package(package_url, work_dir, progress_callback)
    if progress_callback:
        progress_callback(55)

    source_root = extract_update_package(package_path, work_dir, progress_callback)
    helper_path = create_update_helper(
        source_root=source_root,
        target_root=target_root,
        restart_command=restart_command,
        work_dir=work_dir,
    )
    if progress_callback:
        progress_callback(100)
    return helper_path


def launch_helper_and_exit(helper_path: Path) -> None:
    os.startfile(str(helper_path))
