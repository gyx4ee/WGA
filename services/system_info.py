from __future__ import annotations

import platform
import shutil
import subprocess
from dataclasses import dataclass

try:
    import psutil
except Exception:  # pragma: no cover - app still starts with fallbacks
    psutil = None


@dataclass
class UsageSnapshot:
    cpu: int
    ram: int
    disk: int
    disk_detail: str
    gpu: str


@dataclass
class DiskSnapshot:
    letter: str
    percent: int
    used_gb: float
    total_gb: float
    free_gb: float
    filesystem: str


def usage_snapshot() -> UsageSnapshot:
    cpu = int(psutil.cpu_percent(interval=None)) if psutil else 0
    ram = int(psutil.virtual_memory().percent) if psutil else 0
    disks = disk_snapshots()
    disk = _combined_disk_percent(disks)
    disk_detail = _disk_detail(disks)
    return UsageSnapshot(cpu=cpu, ram=ram, disk=disk, disk_detail=disk_detail, gpu=_gpu_name())


def static_system_rows() -> list[tuple[str, str]]:
    rows = [
        ("Windows", _windows_edition()),
        ("Версия", platform.version()),
        ("Build", platform.release()),
        ("CPU", platform.processor() or "Unknown"),
        ("RAM", _ram_total()),
        ("Motherboard", _wmic_value("baseboard", "Product") or "Unknown"),
        ("GPU", _gpu_name()),
    ]
    for disk in disk_snapshots():
        rows.append((f"Диск {disk.letter}", f"{disk.used_gb:.0f}/{disk.total_gb:.0f} GB ({disk.percent}%), свободни {disk.free_gb:.0f} GB, {disk.filesystem}"))
    return rows


def disk_snapshots() -> list[DiskSnapshot]:
    if not psutil:
        return []
    disks: list[DiskSnapshot] = []
    seen: set[str] = set()
    for partition in psutil.disk_partitions(all=False):
        mountpoint = partition.mountpoint
        if not mountpoint:
            continue
        letter = partition.device.rstrip("\\").rstrip(":").upper()
        if not letter or letter in seen:
            continue
        opts = partition.opts.lower()
        if "cdrom" in opts:
            continue
        try:
            usage = psutil.disk_usage(mountpoint)
        except (PermissionError, OSError):
            continue
        if usage.total <= 0:
            continue
        seen.add(letter)
        disks.append(
            DiskSnapshot(
                letter=f"{letter}:",
                percent=int(usage.percent),
                used_gb=usage.used / (1024 ** 3),
                total_gb=usage.total / (1024 ** 3),
                free_gb=usage.free / (1024 ** 3),
                filesystem=partition.fstype or "FS?",
            )
        )
    return disks


def _combined_disk_percent(disks: list[DiskSnapshot]) -> int:
    total = sum(disk.total_gb for disk in disks)
    used = sum(disk.used_gb for disk in disks)
    return int((used / total) * 100) if total else 0


def _disk_detail(disks: list[DiskSnapshot]) -> str:
    if not disks:
        return "Няма намерени дискове"
    return " | ".join(f"{disk.letter} {disk.percent}%" for disk in disks)


def _ram_total() -> str:
    if not psutil:
        return "Unknown"
    gb = psutil.virtual_memory().total / (1024 ** 3)
    return f"{gb:.1f} GB"


def _windows_edition() -> str:
    caption = _wmic_value("os", "Caption")
    return caption or platform.platform()


def _gpu_name() -> str:
    name = _wmic_value("path win32_VideoController", "Name")
    return name or "GPU N/A"


def _wmic_value(alias: str, field: str) -> str:
    if not shutil.which("wmic"):
        return ""
    try:
        result = subprocess.run(
            ["wmic", *alias.split(), "get", field, "/value"],
            capture_output=True,
            text=True,
            timeout=4,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        return ""
    for line in result.stdout.splitlines():
        if line.lower().startswith(field.lower() + "="):
            return line.split("=", 1)[1].strip()
    return ""
