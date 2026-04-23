from __future__ import annotations

import ctypes
from dataclasses import dataclass
from pathlib import Path


DRIVE_REMOVABLE = 2
DRIVE_FIXED = 3


@dataclass(frozen=True)
class RuntimeStorageInfo:
    drive: str
    drive_letter: str
    drive_type: str
    drive_type_label: str
    is_usb: bool
    is_fixed: bool
    program_path: str
    installers_root: Path
    installers_available: bool


def _normalize_drive(path: Path) -> str:
    drive = path.drive or path.anchor.rstrip("\\")
    return drive or ""


def _kernel_drive_type(drive_root: str) -> str:
    result = ctypes.windll.kernel32.GetDriveTypeW(f"{drive_root}\\")
    if result == DRIVE_REMOVABLE:
        return "Removable"
    if result == DRIVE_FIXED:
        return "Fixed"
    return "Unknown"


def detect_drive_type(path: Path) -> str:
    drive_root = _normalize_drive(path)
    if not drive_root:
        return "Unknown"
    return _kernel_drive_type(drive_root)


def describe_drive_type(drive_type: str) -> str:
    if drive_type == "Removable":
        return "USB / Flash Drive"
    if drive_type == "Fixed":
        return "SSD / HDD"
    return "Unknown Drive Type"


def resolve_installers_root(program_root: Path) -> Path:
    drive_root = Path(f"{_normalize_drive(program_root)}\\") if _normalize_drive(program_root) else program_root.anchor
    drive_installers = Path(drive_root) / "Installers" if drive_root else None
    local_installers = program_root / "Installers"
    parent_installers = program_root.parent / "Installers"

    # Portable builds must keep their resources next to WGA.exe. This makes a
    # copied flash-drive folder self-contained and avoids accidentally using C:\Installers.
    candidates = [path for path in (local_installers, drive_installers, parent_installers) if path is not None]

    seen: set[str] = set()
    for candidate in candidates:
        normalized = str(candidate).lower()
        if normalized in seen:
            continue
        seen.add(normalized)
        if candidate.exists():
            return candidate
    return candidates[0]


def ensure_installers_root(program_root: Path) -> Path:
    installers_root = resolve_installers_root(program_root)
    installers_root.mkdir(parents=True, exist_ok=True)
    return installers_root


def get_runtime_storage_info(program_root: Path) -> RuntimeStorageInfo:
    drive = _normalize_drive(program_root)
    drive_letter = drive[:1].upper() if drive else ""
    drive_type = detect_drive_type(program_root)
    installers_root = ensure_installers_root(program_root)
    return RuntimeStorageInfo(
        drive=drive or "Unknown",
        drive_letter=drive_letter,
        drive_type=drive_type,
        drive_type_label=describe_drive_type(drive_type),
        is_usb=drive_type == "Removable",
        is_fixed=drive_type == "Fixed",
        program_path=str(program_root),
        installers_root=installers_root,
        installers_available=installers_root.exists(),
    )
