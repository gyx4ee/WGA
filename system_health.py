from __future__ import annotations

import ctypes
import json
import os
import math
import platform
import shutil
import socket
import subprocess
from dataclasses import dataclass
from pathlib import Path


POWERSHELL = [
    "powershell",
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-Command",
]


MEMORY_TYPE_MAP = {
    20: "DDR",
    21: "DDR2",
    22: "DDR2 FB-DIMM",
    24: "DDR3",
    26: "DDR4",
    34: "DDR5",
}


@dataclass
class HealthItem:
    label: str
    value: str
    ok: bool


class MemoryStatusEx(ctypes.Structure):
    _fields_ = [
        ("dwLength", ctypes.c_uint32),
        ("dwMemoryLoad", ctypes.c_uint32),
        ("ullTotalPhys", ctypes.c_uint64),
        ("ullAvailPhys", ctypes.c_uint64),
        ("ullTotalPageFile", ctypes.c_uint64),
        ("ullAvailPageFile", ctypes.c_uint64),
        ("ullTotalVirtual", ctypes.c_uint64),
        ("ullAvailVirtual", ctypes.c_uint64),
        ("ullAvailExtendedVirtual", ctypes.c_uint64),
    ]


def _run_powershell_json(command: str) -> list[dict[str, object]]:
    result = subprocess.run(
        POWERSHELL + [f"{command} | ConvertTo-Json -Depth 4"],
        capture_output=True,
        text=True,
        check=False,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if result.returncode != 0 or not result.stdout.strip():
        return []

    try:
        data = json.loads(result.stdout)
    except json.JSONDecodeError:
        return []

    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        return [data]
    return []


def _bytes_to_gb(value: int) -> float:
    return value / (1024 ** 3)


def _memory_snapshot() -> tuple[float, float, int]:
    status = MemoryStatusEx()
    status.dwLength = ctypes.sizeof(MemoryStatusEx)
    ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status))
    total = _bytes_to_gb(int(status.ullTotalPhys))
    available = _bytes_to_gb(int(status.ullAvailPhys))
    used = max(0.0, total - available)
    return total, used, int(status.dwMemoryLoad)


def _cpu_name() -> str:
    rows = _run_powershell_json("Get-CimInstance Win32_Processor | Select-Object -First 1 Name")
    if rows and rows[0].get("Name"):
        return str(rows[0]["Name"]).strip()
    return platform.processor() or "Unknown CPU"


def _single_cim_value(command: str, key: str) -> str:
    rows = _run_powershell_json(command)
    if rows and rows[0].get(key) not in (None, ""):
        return str(rows[0][key]).strip()
    return "Unknown"


def _os_details() -> str:
    caption = _single_cim_value(
        "Get-CimInstance Win32_OperatingSystem | Select-Object -First 1 Caption",
        "Caption",
    )
    build = _single_cim_value(
        "Get-CimInstance Win32_OperatingSystem | Select-Object -First 1 BuildNumber",
        "BuildNumber",
    )
    if caption == "Unknown":
        caption = f"{platform.system()} {platform.release()}".strip() or "Unknown OS"
    if build == "Unknown":
        build = platform.version() or "Unknown"
    arch = platform.machine() or "Unknown arch"
    return f"{caption} / build {build} / {arch}"


def _computer_identity() -> str:
    computer = os.environ.get("COMPUTERNAME") or socket.gethostname() or "Unknown PC"
    user = os.environ.get("USERNAME") or "Unknown user"
    domain = os.environ.get("USERDOMAIN") or ""
    identity = f"{computer} / {domain}\\{user}" if domain else f"{computer} / {user}"
    return identity


def _primary_ip() -> str:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            return sock.getsockname()[0]
    except OSError:
        try:
            return socket.gethostbyname(socket.gethostname())
        except OSError:
            return "Unavailable"


def _uptime() -> str:
    ticks = ctypes.windll.kernel32.GetTickCount64()
    seconds = int(ticks / 1000)
    days, remainder = divmod(seconds, 86400)
    hours, remainder = divmod(remainder, 3600)
    minutes, _ = divmod(remainder, 60)
    if days:
        return f"{days}d {hours}h {minutes}m"
    return f"{hours}h {minutes}m"


def _bios_version() -> str:
    manufacturer = _single_cim_value(
        "Get-CimInstance Win32_BIOS | Select-Object -First 1 Manufacturer",
        "Manufacturer",
    )
    version = _single_cim_value(
        "Get-CimInstance Win32_BIOS | Select-Object -First 1 SMBIOSBIOSVersion",
        "SMBIOSBIOSVersion",
    )
    return f"{manufacturer} {version}".strip()


def _motherboard() -> str:
    rows = _run_powershell_json(
        "Get-CimInstance Win32_BaseBoard | Select-Object -First 1 Manufacturer, Product"
    )
    if not rows:
        return "Unknown"
    manufacturer = str(rows[0].get("Manufacturer") or "").strip()
    product = str(rows[0].get("Product") or "").strip()
    return f"{manufacturer} {product}".strip() or "Unknown"


def _gpu_summary() -> str:
    rows = _run_powershell_json(
        "Get-CimInstance Win32_VideoController | Select-Object Name, AdapterRAM"
    )
    if not rows:
        return "Unknown GPU"
    values: list[str] = []
    for row in rows[:2]:
        name = str(row.get("Name") or "GPU").strip()
        ram = row.get("AdapterRAM")
        if isinstance(ram, (int, float)) and ram > 0:
            values.append(f"{name} ({_bytes_to_gb(int(ram)):.1f} GB)")
        else:
            values.append(name)
    return " / ".join(values)


def _battery_status() -> tuple[str, bool]:
    rows = _run_powershell_json(
        "Get-CimInstance Win32_Battery | Select-Object -First 1 EstimatedChargeRemaining, BatteryStatus"
    )
    if not rows:
        return "Desktop / no battery detected", True
    charge = rows[0].get("EstimatedChargeRemaining")
    status = rows[0].get("BatteryStatus")
    if isinstance(charge, (int, float)):
        ok = int(charge) >= 20
        return f"{int(charge)}% / status {status or 'N/A'}", ok
    return "Battery detected / status unknown", False


def _secure_boot_status() -> tuple[str, bool]:
    result = subprocess.run(
        POWERSHELL + ["Confirm-SecureBootUEFI"],
        capture_output=True,
        text=True,
        check=False,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    value = result.stdout.strip().lower()
    if value == "true":
        return "Enabled", True
    if value == "false":
        return "Disabled", False
    return "Unavailable", False


def _ram_type() -> str:
    rows = _run_powershell_json("Get-CimInstance Win32_PhysicalMemory | Select-Object SMBIOSMemoryType, MemoryType, Speed")
    if not rows:
        return "Unknown"

    ram_types: list[str] = []
    speeds: list[str] = []
    for row in rows:
        type_code = row.get("SMBIOSMemoryType") or row.get("MemoryType")
        if isinstance(type_code, (int, float)):
            mapped = MEMORY_TYPE_MAP.get(int(type_code))
            if mapped and mapped not in ram_types:
                ram_types.append(mapped)
        speed = row.get("Speed")
        if isinstance(speed, (int, float)) and int(speed) > 0:
            speeds.append(f"{int(speed)} MHz")

    ram_type = ", ".join(ram_types) if ram_types else "Unknown"
    speed_text = speeds[0] if speeds else "Speed N/A"
    return f"{ram_type} / {speed_text}"


def _cpu_voltage() -> tuple[str, bool]:
    rows = _run_powershell_json("Get-CimInstance Win32_Processor | Select-Object -First 1 CurrentVoltage")
    if not rows or rows[0].get("CurrentVoltage") in (None, 0, ""):
        return "Sensor unavailable", False

    raw_value = int(rows[0]["CurrentVoltage"])
    voltage = raw_value / 10 if raw_value > 10 else float(raw_value)
    ok = 0.7 <= voltage <= 1.5
    return f"{voltage:.2f} V", ok


def _temperature() -> tuple[str, bool]:
    rows = _run_powershell_json(
        r"Get-CimInstance MSAcpi_ThermalZoneTemperature -Namespace root/wmi | Select-Object CurrentTemperature"
    )
    values: list[float] = []
    for row in rows:
        raw = row.get("CurrentTemperature")
        if isinstance(raw, (int, float)) and raw > 0:
            celsius = (float(raw) / 10) - 273.15
            if -20 < celsius < 150:
                values.append(celsius)

    if not values:
        return "Sensor unavailable", False

    highest = max(values)
    return f"{highest:.1f} C", highest < 85


def _disk_items() -> list[HealthItem]:
    items: list[HealthItem] = []
    volume_rows = _run_powershell_json(
        "Get-CimInstance Win32_LogicalDisk | Select-Object DeviceID, FileSystem, VolumeName"
    )
    volume_info = {
        str(row.get("DeviceID") or "").rstrip(":").upper(): row
        for row in volume_rows
    }
    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        drive = Path(f"{letter}:/")
        if not drive.exists():
            continue
        try:
            usage = shutil.disk_usage(drive)
        except OSError:
            continue

        total_gb = _bytes_to_gb(usage.total)
        used_gb = _bytes_to_gb(usage.used)
        percent = math.floor((usage.used / usage.total) * 100) if usage.total else 0
        free_gb = _bytes_to_gb(usage.free)
        row = volume_info.get(letter, {})
        filesystem = str(row.get("FileSystem") or "FS?").strip()
        volume_name = str(row.get("VolumeName") or "").strip()
        name_part = f" {volume_name}" if volume_name else ""
        ok = percent < 90
        items.append(
            HealthItem(
                label=f"Disk {letter}:",
                value=f"{used_gb:.0f}/{total_gb:.0f} GB ({percent}%), free {free_gb:.0f} GB, {filesystem}{name_part}",
                ok=ok,
            )
        )
    return items or [HealthItem(label="Disk:", value="No drives detected", ok=False)]


def collect_health_items() -> list[HealthItem]:
    total_ram, used_ram, ram_load = _memory_snapshot()
    ram_ok = ram_load < 85
    temperature_value, temperature_ok = _temperature()
    voltage_value, voltage_ok = _cpu_voltage()
    ram_type_value = _ram_type()
    battery_value, battery_ok = _battery_status()
    secure_boot_value, secure_boot_ok = _secure_boot_status()

    items = [
        HealthItem(label="OS:", value=_os_details(), ok=True),
        HealthItem(label="PC/User:", value=_computer_identity(), ok=True),
        HealthItem(label="IP:", value=_primary_ip(), ok=True),
        HealthItem(label="Uptime:", value=_uptime(), ok=True),
        HealthItem(label="CPU:", value=_cpu_name(), ok=True),
        HealthItem(label="Temperature:", value=temperature_value, ok=temperature_ok),
        HealthItem(label="CPU Voltage:", value=voltage_value, ok=voltage_ok),
        HealthItem(label="GPU:", value=_gpu_summary(), ok=True),
        HealthItem(
            label="RAM:",
            value=f"{used_ram:.1f}/{total_ram:.1f} GB ({ram_load}%)",
            ok=ram_ok,
        ),
        HealthItem(label="RAM Type:", value=ram_type_value, ok="Unknown" not in ram_type_value),
        HealthItem(label="Motherboard:", value=_motherboard(), ok=True),
        HealthItem(label="BIOS:", value=_bios_version(), ok=True),
        HealthItem(label="Secure Boot:", value=secure_boot_value, ok=secure_boot_ok),
        HealthItem(label="Battery:", value=battery_value, ok=battery_ok),
    ]
    items.extend(_disk_items())
    return items
