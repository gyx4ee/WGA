from __future__ import annotations

import ctypes
import shutil
import subprocess
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


DRIVE_REMOVABLE = 2


@dataclass
class DriverBackupResult:
    backup_dir: Path
    zip_path: Path | None
    log_path: Path
    drivers_list_path: Path
    restore_script_path: Path


def desktop_path() -> Path:
    return Path.home() / "Desktop"


def onedrive_path() -> Path | None:
    candidates = [
        Path.home() / "OneDrive",
        Path.home() / "OneDrive - Personal",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def timestamp_string() -> str:
    return datetime.now().strftime("%Y-%m-%d_%H-%M")


def create_backup_folder(base_path: Path) -> Path:
    destination = base_path / f"DriversBackup_{timestamp_string()}"
    destination.mkdir(parents=True, exist_ok=True)
    return destination


def run_command(command: list[str], cwd: Path | None = None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        capture_output=True,
        text=True,
        check=False,
        cwd=str(cwd) if cwd else None,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def export_drivers(destination: Path, mode: str) -> tuple[subprocess.CompletedProcess[str], Path]:
    log_path = destination / "backup_log.txt"
    if mode == "clean":
        result = run_command(["pnputil", "/export-driver", "*", str(destination)])
    elif mode == "full":
        result = run_command(["dism", "/online", "/export-driver", f"/destination:{destination}"])
    else:
        raise ValueError(f"Unsupported backup mode: {mode}")

    log_text = "\n".join(part.strip() for part in (result.stdout, result.stderr) if part and part.strip())
    log_path.write_text(log_text or "No output captured.", encoding="utf-8")
    return result, log_path


def create_driver_list(destination: Path) -> Path:
    drivers_list_path = destination / "drivers_list.txt"
    result = run_command(["pnputil", "/enum-drivers"])
    content = "\n".join(part.strip() for part in (result.stdout, result.stderr) if part and part.strip())
    drivers_list_path.write_text(content or "No driver list output captured.", encoding="utf-8")
    return drivers_list_path


def create_restore_script(destination: Path) -> Path:
    restore_script_path = destination / "RESTORE_DRIVERS.bat"
    restore_script_path.write_text(
        "@echo off\n"
        "echo ======================================\n"
        "echo      DRIVER RESTORE TOOL\n"
        "echo ======================================\n"
        "echo Installing all drivers...\n"
        f'pnputil /add-driver "{destination}\\*.inf" /subdirs /install\n'
        "echo DONE!\n"
        "pause\n",
        encoding="utf-8",
    )
    return restore_script_path


def compress_backup(destination: Path, delete_original: bool = False) -> Path:
    archive_path = Path(shutil.make_archive(str(destination), "zip", root_dir=destination.parent, base_dir=destination.name))
    if delete_original:
        shutil.rmtree(destination, ignore_errors=True)
    return archive_path


def detect_removable_drives() -> list[Path]:
    drives: list[Path] = []
    kernel32 = ctypes.windll.kernel32
    for letter in "DEFGHIJKLMNOPQRSTUVWXYZ":
        root = Path(f"{letter}:\\")
        if not root.exists():
            continue
        if kernel32.GetDriveTypeW(str(root)) == DRIVE_REMOVABLE:
            drives.append(root)
    return drives


def create_recovery_usb(last_backup_dir: Path, usb_root: Path) -> tuple[Path, Path]:
    recovery_dir = usb_root / "DriverRecoveryBackup"
    if recovery_dir.exists():
        shutil.rmtree(recovery_dir, ignore_errors=True)
    shutil.copytree(last_backup_dir, recovery_dir)
    restore_script = usb_root / "RESTORE_DRIVERS.bat"
    restore_script.write_text(
        "@echo off\n"
        "echo ======================================\n"
        "echo      DRIVER RESTORE TOOL\n"
        "echo ======================================\n"
        "echo Installing all drivers...\n"
        f'pnputil /add-driver "{recovery_dir}\\*.inf" /subdirs /install\n'
        "echo DONE!\n"
        "pause\n",
        encoding="utf-8",
    )
    return recovery_dir, restore_script


def restore_drivers_from_backup(backup_dir: Path) -> subprocess.CompletedProcess[str]:
    return run_command(["pnputil", "/add-driver", str(backup_dir / "*.inf"), "/subdirs", "/install"])


def generate_pc_report(destination: Path) -> Path:
    report_path = destination / "PC_Report.txt"
    sections = [
        ("CPU INFO", "Get-CimInstance Win32_Processor | Select-Object Name,NumberOfCores,NumberOfLogicalProcessors | Format-List"),
        ("RAM INFO", "Get-CimInstance Win32_PhysicalMemory | Select-Object Capacity,Speed,Manufacturer,PartNumber | Format-Table -AutoSize"),
        ("GPU INFO", "Get-CimInstance Win32_VideoController | Select-Object Name,AdapterRAM,DriverVersion | Format-Table -AutoSize"),
        ("MOTHERBOARD INFO", "Get-CimInstance Win32_BaseBoard | Select-Object Manufacturer,Product,SerialNumber | Format-List"),
        ("BIOS INFO", "Get-CimInstance Win32_BIOS | Select-Object SMBIOSBIOSVersion,SerialNumber,ReleaseDate | Format-List"),
        ("DISK DRIVES", "Get-CimInstance Win32_DiskDrive | Select-Object Model,Size,MediaType | Format-Table -AutoSize"),
        ("OPERATING SYSTEM", "Get-CimInstance Win32_OperatingSystem | Select-Object Caption,Version,BuildNumber,OSArchitecture | Format-List"),
        ("NETWORK ADAPTERS", "Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object {$_.IPEnabled} | Select-Object Description,IPAddress,MACAddress | Format-List"),
    ]
    lines = [f"Generated: {datetime.now().isoformat(sep=' ', timespec='seconds')}", ""]
    for title, script in sections:
        lines.append(f"=== {title} ===")
        result = run_command(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script])
        content = "\n".join(part.strip() for part in (result.stdout, result.stderr) if part and part.strip())
        lines.append(content or "No data returned.")
        lines.append("")
    report_path.write_text("\n".join(lines), encoding="utf-8")
    return report_path
