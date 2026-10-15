import ctypes
import json
import platform
import shutil
import socket
import subprocess
import sys
import uuid
from datetime import UTC, datetime
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit
from urllib.request import Request, urlopen
import tempfile
import time
import zipfile


APP_NAME = "WinSys Guardian Agent"
STATUS_TASK_NAME = "WGA Agent Status"
STARTUP_TASK_NAME = "WGA Agent Startup"
INSTALL_DIR = Path("C:/ProgramData/WGA/Agent")
INSTALLED_AGENT_PATH = INSTALL_DIR / "wga_agent.exe"
INSTALLED_SCRIPT_PATH = INSTALL_DIR / "wga_agent.py"
INSTALLED_VERSION_PATH = INSTALL_DIR / "version.json"
ADMIN_SHORTCUT_PATH = Path("C:/ProgramData/Microsoft/Windows/Start Menu/Programs/WGA Agent Admin Console.lnk")
LOG_PATH = INSTALL_DIR / "wga_agent.log"
DEFAULT_STATUS_PATH = INSTALL_DIR / "wga_agent_status.json"
REALTIME_ACTIVITY_PATH = INSTALL_DIR / "wga_realtime_activity.json"
REQUEST_DIR = INSTALL_DIR / "requests"
COMMAND_QUEUE_PATH = REQUEST_DIR / "commands.json"
COMMAND_RESULT_PATH = REQUEST_DIR / "command_results.json"
DEFAULT_INTERVAL_MINUTES = 5
DEFAULT_VERSION_INFO_URL = "https://raw.githubusercontent.com/gyx4ee/WGA/refs/heads/main/version.json"
PROTECTED_PROCESS_NAMES = {
    "csrss",
    "dwm",
    "explorer",
    "lsass",
    "services",
    "smss",
    "svchost",
    "system",
    "wininit",
    "winlogon",
    "wga_agent",
}
RDP_PORT = 3389


def is_online(timeout: float = 2.0) -> bool:
    try:
        with socket.create_connection(("8.8.8.8", 53), timeout=timeout):
            return True
    except OSError:
        return False


def get_local_ip_addresses() -> list[str]:
    addresses: set[str] = set()
    try:
        for family, _, _, _, sockaddr in socket.getaddrinfo(socket.gethostname(), None):
            if family == socket.AF_INET:
                ip = sockaddr[0]
                if not ip.startswith("127."):
                    addresses.add(ip)
    except OSError:
        pass

    if not addresses:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
                sock.connect(("8.8.8.8", 80))
                addresses.add(sock.getsockname()[0])
        except OSError:
            pass

    return sorted(addresses)


def format_duration(seconds: int) -> str:
    days, seconds = divmod(seconds, 86400)
    hours, seconds = divmod(seconds, 3600)
    minutes, seconds = divmod(seconds, 60)
    parts = []
    if days:
        parts.append(f"{days}d")
    if hours:
        parts.append(f"{hours}h")
    if minutes:
        parts.append(f"{minutes}m")
    if seconds or not parts:
        parts.append(f"{seconds}s")
    return " ".join(parts)


def bytes_to_gb(bytes_count: int) -> float:
    return round(bytes_count / 1024 / 1024 / 1024, 2)


def run_powershell_json(script: str, timeout: int = 12) -> object | None:
    try:
        output = subprocess.check_output(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                script,
            ],
            text=True,
            stderr=subprocess.DEVNULL,
            timeout=timeout,
        )
        if not output.strip():
            return None
        return json.loads(output)
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired, json.JSONDecodeError):
        return None


def normalized_list(value: object | None) -> list[object]:
    if value is None:
        return []
    return value if isinstance(value, list) else [value]


def get_system_uptime_seconds() -> int:
    if platform.system() == "Windows":
        try:
            return int(ctypes.windll.kernel32.GetTickCount64() / 1000)
        except AttributeError:
            return int(ctypes.windll.kernel32.GetTickCount() / 1000)
    return 0


def get_windows_update_status() -> dict[str, object]:
    status = {
        "up_to_date": True,
        "pending_updates": 0,
        "error": None,
    }
    if platform.system() != "Windows":
        return status

    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        "& { $u=(New-Object -ComObject Microsoft.Update.Session).CreateUpdateSearcher().Search('IsInstalled=0'); $u.Updates.Count }",
    ]

    try:
        output = subprocess.check_output(command, text=True, stderr=subprocess.DEVNULL)
        count = int(output.strip() or "0")
        status["pending_updates"] = count
        status["up_to_date"] = count == 0
    except Exception as exc:
        status["up_to_date"] = False
        status["error"] = str(exc)

    return status


def get_windows_hardware_info() -> dict[str, object]:
    info: dict[str, object] = {}
    if platform.system() != "Windows":
        return info

    system_data = run_powershell_json(
        "Get-CimInstance Win32_ComputerSystem | Select-Object Name,Manufacturer,Model,TotalPhysicalMemory | ConvertTo-Json -Compress"
    )
    if system_data:
        sys_info = system_data[0] if isinstance(system_data, list) else system_data
        total_physical_memory = int(sys_info.get("TotalPhysicalMemory", 0) or 0)
        info["computer_system"] = {
            "name": sys_info.get("Name", "N/A"),
            "manufacturer": sys_info.get("Manufacturer", "N/A"),
            "model": sys_info.get("Model", "N/A"),
            "total_physical_memory_bytes": total_physical_memory,
            "total_physical_memory_gb": bytes_to_gb(total_physical_memory),
        }

    memory_data = run_powershell_json(
        "Get-CimInstance Win32_OperatingSystem | Select-Object TotalVisibleMemorySize,FreePhysicalMemory,TotalVirtualMemorySize,FreeVirtualMemory | ConvertTo-Json -Compress"
    )
    if memory_data:
        mem_info = memory_data[0] if isinstance(memory_data, list) else memory_data
        total_ram = int(mem_info.get("TotalVisibleMemorySize", 0) or 0) * 1024
        free_ram = int(mem_info.get("FreePhysicalMemory", 0) or 0) * 1024
        info["memory"] = {
            "total_bytes": total_ram,
            "free_bytes": free_ram,
            "total_gb": bytes_to_gb(total_ram),
            "free_gb": bytes_to_gb(free_ram),
            "used_gb": bytes_to_gb(total_ram - free_ram),
            "used_percent": round((total_ram - free_ram) / total_ram * 100, 1) if total_ram else 0,
        }

    cpu_data = run_powershell_json(
        "Get-CimInstance Win32_Processor | Select-Object Name,Manufacturer,NumberOfCores,NumberOfLogicalProcessors,MaxClockSpeed,LoadPercentage,Architecture | ConvertTo-Json -Compress"
    )
    info["cpu"] = []
    for cpu in normalized_list(cpu_data):
        info["cpu"].append({
            "name": cpu.get("Name", "N/A"),
            "manufacturer": cpu.get("Manufacturer", "N/A"),
            "cores": int(cpu.get("NumberOfCores", 0) or 0),
            "logical_processors": int(cpu.get("NumberOfLogicalProcessors", 0) or 0),
            "max_clock_mhz": int(cpu.get("MaxClockSpeed", 0) or 0),
            "load_percent": int(cpu.get("LoadPercentage", 0) or 0),
            "architecture": int(cpu.get("Architecture", 0) or 0),
        })

    gpu_data = run_powershell_json(
        "Get-CimInstance Win32_VideoController | Select-Object Name,DriverVersion,VideoProcessor,AdapterRAM | ConvertTo-Json -Compress"
    )
    info["video_controllers"] = []
    for gpu in normalized_list(gpu_data):
        adapter_ram = int(gpu.get("AdapterRAM", 0) or 0)
        info["video_controllers"].append({
            "name": gpu.get("Name", "N/A"),
            "driver_version": gpu.get("DriverVersion", "N/A"),
            "video_processor": gpu.get("VideoProcessor", "N/A"),
            "adapter_ram_bytes": adapter_ram,
            "adapter_ram_gb": bytes_to_gb(adapter_ram),
        })

    disk_data = run_powershell_json(
        "Get-CimInstance Win32_LogicalDisk -Filter \"DriveType=3\" | Select-Object DeviceID,FileSystem,Size,FreeSpace | ConvertTo-Json -Compress"
    )
    info["disks"] = []
    for disk in normalized_list(disk_data):
        size = int(disk.get("Size", 0) or 0)
        free = int(disk.get("FreeSpace", 0) or 0)
        info["disks"].append({
            "device_id": disk.get("DeviceID", "N/A"),
            "filesystem": disk.get("FileSystem", "N/A"),
            "size_bytes": size,
            "free_bytes": free,
            "size_gb": bytes_to_gb(size),
            "free_gb": bytes_to_gb(free),
            "used_percent": round((size - free) / size * 100, 1) if size else 0,
        })

    temp_data = run_powershell_json(
        "Get-CimInstance -Namespace root/wmi -ClassName MSAcpi_ThermalZoneTemperature | Select-Object InstanceName,CurrentTemperature,Active | ConvertTo-Json -Compress"
    )
    info["thermal_zones"] = []
    for temp in normalized_list(temp_data):
        temp_raw = int(temp.get("CurrentTemperature", 0) or 0)
        temp_c = round((temp_raw / 10.0) - 273.15, 1) if temp_raw else None
        info["thermal_zones"].append({
            "instance_name": temp.get("InstanceName", "N/A"),
            "current_temperature_c": temp_c,
            "active": temp.get("Active", False),
        })

    perf_data = run_powershell_json(
        "Get-CimInstance Win32_PerfFormattedData_PerfOS_Processor | Where-Object { $_.Name -eq '_Total' } | Select-Object PercentProcessorTime | ConvertTo-Json -Compress"
    )
    if perf_data:
        perf = perf_data[0] if isinstance(perf_data, list) else perf_data
        info["cpu_total_load_percent"] = int(perf.get("PercentProcessorTime", 0) or 0)

    return info


def collect_system_info() -> dict[str, object]:
    now = datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    uptime_seconds = get_system_uptime_seconds()
    return {
        "timestamp": now,
        "hostname": socket.gethostname(),
        "node": platform.node(),
        "platform": platform.system(),
        "platform_release": platform.release(),
        "platform_version": platform.version(),
        "machine": platform.machine(),
        "processor": platform.processor(),
        "python_version": platform.python_version(),
        "online": is_online(),
        "local_ips": get_local_ip_addresses(),
        "system_uptime_seconds": uptime_seconds,
        "system_uptime": format_duration(uptime_seconds),
        "update_status": get_windows_update_status(),
        "hardware_info": get_windows_hardware_info(),
        "realtime_activity": collect_realtime_activity(),
        "uuid": str(uuid.uuid4()),
    }


def get_foreground_window_info() -> dict[str, object]:
    if platform.system() != "Windows":
        return {}
    try:
        user32 = ctypes.windll.user32
        hwnd = user32.GetForegroundWindow()
        if not hwnd:
            return {}
        length = user32.GetWindowTextLengthW(hwnd)
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buffer, length + 1)
        process_id = ctypes.c_ulong()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
        process_name = ""
        process_data = run_powershell_json(
            f"Get-Process -Id {int(process_id.value)} -ErrorAction SilentlyContinue | "
            "Select-Object Id,ProcessName,Path | ConvertTo-Json -Compress",
            timeout=3,
        )
        if isinstance(process_data, dict):
            process_name = str(process_data.get("ProcessName", ""))
        return {
            "window_title": buffer.value,
            "process_id": int(process_id.value),
            "process_name": process_name,
        }
    except Exception as exc:
        return {"error": str(exc)}


def get_interactive_user() -> str:
    if platform.system() != "Windows":
        return ""
    data = run_powershell_json(
        "Get-CimInstance Win32_ComputerSystem | Select-Object UserName | ConvertTo-Json -Compress",
        timeout=4,
    )
    if isinstance(data, dict):
        return str(data.get("UserName", "") or "")
    return ""


def get_running_process_snapshot(limit: int = 30) -> list[dict[str, object]]:
    if platform.system() != "Windows":
        return []
    script = f"""
Get-Process |
Sort-Object CPU -Descending |
Select-Object -First {limit} Id,ProcessName,CPU,WorkingSet64,MainWindowTitle |
ConvertTo-Json -Compress
"""
    data = run_powershell_json(script, timeout=4)
    processes = []
    for item in normalized_list(data):
        if not isinstance(item, dict):
            continue
        processes.append(
            {
                "pid": int(item.get("Id", 0) or 0),
                "name": item.get("ProcessName", ""),
                "cpu_seconds": round(float(item.get("CPU", 0) or 0), 2),
                "memory_mb": round(int(item.get("WorkingSet64", 0) or 0) / 1024 / 1024, 1),
                "window_title": item.get("MainWindowTitle", ""),
            }
        )
    return processes


def get_hung_process_snapshot(limit: int = 50) -> list[dict[str, object]]:
    if platform.system() != "Windows":
        return []
    script = f"""
Get-Process |
Where-Object {{ $_.MainWindowHandle -ne 0 -and $_.Responding -eq $false }} |
Select-Object -First {limit} Id,ProcessName,MainWindowTitle,Path |
ConvertTo-Json -Compress
"""
    data = run_powershell_json(script, timeout=4)
    hung_processes = []
    for item in normalized_list(data):
        if not isinstance(item, dict):
            continue
        hung_processes.append(
            {
                "pid": int(item.get("Id", 0) or 0),
                "name": item.get("ProcessName", ""),
                "window_title": item.get("MainWindowTitle", ""),
                "path": item.get("Path", ""),
            }
        )
    return hung_processes


def get_network_connection_snapshot(limit: int = 80) -> list[dict[str, object]]:
    if platform.system() != "Windows":
        return []
    script = f"""
$connections = Get-NetTCPConnection -State Established -ErrorAction SilentlyContinue |
    Where-Object {{ $_.RemoteAddress -notlike '127.*' -and $_.RemoteAddress -ne '::1' }} |
    Select-Object -First {limit} LocalAddress,LocalPort,RemoteAddress,RemotePort,OwningProcess
$items = foreach ($connection in $connections) {{
    $process = Get-Process -Id $connection.OwningProcess -ErrorAction SilentlyContinue
    [PSCustomObject]@{{
        LocalAddress = $connection.LocalAddress
        LocalPort = $connection.LocalPort
        RemoteAddress = $connection.RemoteAddress
        RemotePort = $connection.RemotePort
        ProcessId = $connection.OwningProcess
        ProcessName = if ($process) {{ $process.ProcessName }} else {{ "" }}
    }}
}}
$items | ConvertTo-Json -Compress
"""
    data = run_powershell_json(script, timeout=5)
    connections = []
    for item in normalized_list(data):
        if not isinstance(item, dict):
            continue
        connections.append(
            {
                "local": f"{item.get('LocalAddress', '')}:{item.get('LocalPort', '')}",
                "remote": f"{item.get('RemoteAddress', '')}:{item.get('RemotePort', '')}",
                "process_id": int(item.get("ProcessId", 0) or 0),
                "process_name": item.get("ProcessName", ""),
            }
        )
    return connections


def get_dns_cache_snapshot(limit: int = 80) -> list[dict[str, object]]:
    if platform.system() != "Windows":
        return []
    script = f"""
Get-DnsClientCache -ErrorAction SilentlyContinue |
Where-Object {{ $_.Entry -and $_.Entry -notlike '*.in-addr.arpa' }} |
Select-Object -First {limit} Entry,Type,Data,TimeToLive |
ConvertTo-Json -Compress
"""
    data = run_powershell_json(script, timeout=4)
    entries = []
    seen: set[str] = set()
    for item in normalized_list(data):
        if not isinstance(item, dict):
            continue
        entry = str(item.get("Entry", "") or "").strip()
        if not entry or entry in seen:
            continue
        seen.add(entry)
        entries.append(
            {
                "domain": entry,
                "type": item.get("Type", ""),
                "data": item.get("Data", ""),
                "ttl": item.get("TimeToLive", ""),
            }
        )
    return entries


def get_rdp_session_snapshot() -> list[dict[str, object]]:
    if platform.system() != "Windows":
        return []
    result = run_hidden(["query", "session"], timeout=8)
    output = result.stdout.strip()
    sessions = []
    for raw_line in output.splitlines()[1:]:
        line = raw_line.strip()
        if not line:
            continue
        active_marker = line.startswith(">")
        if active_marker:
            line = line[1:].strip()
        parts = line.split()
        if len(parts) < 3:
            continue
        session_name = parts[0]
        username = ""
        session_id = None
        state = ""
        if len(parts) >= 4 and parts[2].isdigit():
            username = parts[1]
            session_id = int(parts[2])
            state = parts[3]
        elif parts[1].isdigit():
            session_id = int(parts[1])
            state = parts[2]
        if session_id is None:
            continue
        sessions.append(
            {
                "session_name": session_name,
                "username": username,
                "id": session_id,
                "state": state,
                "current": active_marker,
            }
        )
    return sessions


def get_remote_desktop_status() -> dict[str, object]:
    if platform.system() != "Windows":
        return {"available": False, "reason": "Windows only"}
    deny_data = run_powershell_json(
        "Get-ItemProperty -Path 'HKLM:\\System\\CurrentControlSet\\Control\\Terminal Server' "
        "-Name fDenyTSConnections | Select-Object fDenyTSConnections | ConvertTo-Json -Compress",
        timeout=5,
    )
    nla_data = run_powershell_json(
        "Get-ItemProperty -Path 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Terminal Server\\WinStations\\RDP-Tcp' "
        "-Name UserAuthentication | Select-Object UserAuthentication | ConvertTo-Json -Compress",
        timeout=5,
    )
    service = run_powershell_json(
        "Get-Service TermService | Select-Object Status,StartType | ConvertTo-Json -Compress",
        timeout=5,
    )
    firewall = run_powershell_json(
        "Get-NetFirewallRule -DisplayGroup 'Remote Desktop' -ErrorAction SilentlyContinue | "
        "Where-Object Enabled -eq 'True' | Measure-Object | Select-Object Count | ConvertTo-Json -Compress",
        timeout=5,
    )
    deny = int(deny_data.get("fDenyTSConnections", 1) or 1) if isinstance(deny_data, dict) else 1
    nla = int(nla_data.get("UserAuthentication", 0) or 0) if isinstance(nla_data, dict) else 0
    firewall_count = int(firewall.get("Count", 0) or 0) if isinstance(firewall, dict) else 0
    hostname = socket.gethostname()
    ips = get_local_ip_addresses()
    active_sessions = get_rdp_session_snapshot()
    return {
        "available": True,
        "enabled": deny == 0,
        "nla_enabled": nla == 1,
        "service": service if isinstance(service, dict) else {},
        "firewall_rules_enabled": firewall_count,
        "hostname": hostname,
        "local_ips": ips,
        "port": RDP_PORT,
        "connect_command": f"mstsc /v:{hostname}",
        "connect_command_ip": f"mstsc /v:{ips[0]}" if ips else "",
        "shadow_commands": [
            f"mstsc /shadow:{session.get('id')} /v:{hostname} /control /prompt"
            for session in active_sessions
            if session.get("id") not in {0, None}
        ],
        "sessions": active_sessions,
    }


def enable_remote_desktop() -> str:
    if platform.system() != "Windows":
        raise RuntimeError("Remote Desktop enable is Windows-only.")
    script = r"""
Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections -Value 0
Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name UserAuthentication -Value 1
Set-Service -Name TermService -StartupType Automatic
Start-Service -Name TermService
Enable-NetFirewallRule -DisplayGroup 'Remote Desktop'
"""
    result = run_hidden(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        timeout=60,
    )
    output = f"{result.stdout}\n{result.stderr}".strip()
    if result.returncode != 0:
        raise RuntimeError(output or "Remote Desktop enable failed.")
    return "Remote Desktop е включен с NLA и firewall rules."


def collect_realtime_activity() -> dict[str, object]:
    now = datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return {
        "timestamp": now,
        "interactive_user": get_interactive_user(),
        "foreground_window": get_foreground_window_info(),
        "top_processes": get_running_process_snapshot(),
        "hung_processes": get_hung_process_snapshot(),
        "network_connections": get_network_connection_snapshot(),
        "dns_cache": get_dns_cache_snapshot(),
        "remote_desktop": get_remote_desktop_status(),
    }


def save_realtime_activity(output_path: Path | str = REALTIME_ACTIVITY_PATH) -> Path:
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    data = collect_realtime_activity()
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    return path


def save_status(output_path: Path | str = "wga_agent_status.json") -> Path:
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    data = collect_system_info()
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    return path


def append_command_result(command_id: str, status: str, message: str) -> None:
    results = read_json_file(COMMAND_RESULT_PATH)
    entries = results.get("results", [])
    if not isinstance(entries, list):
        entries = []
    entries.append(
        {
            "id": command_id,
            "status": status,
            "message": message,
            "timestamp": datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
        }
    )
    COMMAND_RESULT_PATH.parent.mkdir(parents=True, exist_ok=True)
    COMMAND_RESULT_PATH.write_text(json.dumps({"results": entries[-200:]}, indent=2), encoding="utf-8")


def is_protected_process_name(process_name: str) -> bool:
    return process_name.strip().lower().removesuffix(".exe") in PROTECTED_PROCESS_NAMES


def kill_hung_process(command: dict[str, object]) -> str:
    raw_pid = command.get("pid")
    raw_name = str(command.get("process_name", "") or command.get("name", "")).strip()
    target_pid = int(raw_pid) if str(raw_pid or "").isdigit() else None
    target_name = raw_name.removesuffix(".exe")
    if target_pid is None and not target_name:
        raise RuntimeError("Липсва pid или process_name за kill_hung_process.")
    if target_name and is_protected_process_name(target_name):
        raise RuntimeError(f"Процесът '{target_name}' е защитен и няма да бъде спиран.")

    hung_processes = get_hung_process_snapshot(limit=200)
    matches = []
    for process in hung_processes:
        process_pid = int(process.get("pid", 0) or 0)
        process_name = str(process.get("name", "") or "")
        if target_pid is not None and process_pid == target_pid:
            matches.append(process)
        elif target_name and process_name.lower() == target_name.lower():
            matches.append(process)

    if not matches:
        raise RuntimeError("Не е намерен забил matching процес. Kill се отказва.")

    stopped = []
    for process in matches[:10]:
        process_pid = int(process.get("pid", 0) or 0)
        process_name = str(process.get("name", "") or "")
        if not process_pid or is_protected_process_name(process_name):
            continue
        result = run_hidden(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                f"Stop-Process -Id {process_pid} -Force -ErrorAction Stop",
            ],
            timeout=20,
        )
        output = f"{result.stdout}\n{result.stderr}".strip()
        if result.returncode != 0:
            raise RuntimeError(output or f"Stop-Process failed for pid {process_pid}.")
        stopped.append(f"{process_name}({process_pid})")

    if not stopped:
        raise RuntimeError("Няма позволен забил процес за спиране.")
    return "Спрени забили процеси: " + ", ".join(stopped)


def process_command_queue() -> None:
    queue = read_json_file(COMMAND_QUEUE_PATH)
    commands = queue.get("commands", [])
    if not isinstance(commands, list) or not commands:
        return

    pending_commands = []
    for item in commands:
        if not isinstance(item, dict):
            continue
        command_id = str(item.get("id", uuid.uuid4())).strip()
        command_type = str(item.get("type", "")).strip()
        if command_type == "status":
            save_status(DEFAULT_STATUS_PATH)
            append_command_result(command_id, "success", "Status snapshot updated.")
        elif command_type == "kill_hung_process":
            try:
                message = kill_hung_process(item)
                append_command_result(command_id, "success", message)
            except Exception as exc:
                append_command_result(command_id, "error", str(exc))
        elif command_type == "enable_remote_desktop":
            try:
                message = enable_remote_desktop()
                append_command_result(command_id, "success", message)
            except Exception as exc:
                append_command_result(command_id, "error", str(exc))
        else:
            pending_commands.append(item)
            append_command_result(
                command_id,
                "deferred",
                f"Command type '{command_type}' is not implemented yet. Future installs will use an allowlist.",
            )

    COMMAND_QUEUE_PATH.write_text(json.dumps({"commands": pending_commands}, indent=2), encoding="utf-8")


def is_windows_admin() -> bool:
    if platform.system() != "Windows":
        return False
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except OSError:
        return False


def run_hidden(command: list[str], timeout: int = 60) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        capture_output=True,
        text=True,
        check=False,
        timeout=timeout,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def log_event(message: str) -> None:
    try:
        INSTALL_DIR.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")
        with LOG_PATH.open("a", encoding="utf-8") as handle:
            handle.write(f"[{timestamp}] {message}\n")
    except OSError:
        pass


def bundled_file(relative_path: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative_path


def read_json_file(path: Path) -> dict[str, object]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def local_version_info() -> dict[str, object]:
    installed = read_json_file(INSTALLED_VERSION_PATH)
    if installed:
        return installed
    bundled = read_json_file(bundled_file("version.json"))
    if bundled:
        return bundled
    project = read_json_file(Path(__file__).resolve().parent / "version.json")
    if project:
        return project
    return {"version": "0.0.0", "version_info_url": DEFAULT_VERSION_INFO_URL}


def normalize_version(version: str) -> tuple[int, ...]:
    parts: list[int] = []
    for item in version.strip().split("."):
        try:
            parts.append(int(item))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def with_cache_buster(url: str) -> str:
    parts = urlsplit(url)
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    query["_wga_agent_ts"] = str(int(time.time()))
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query), parts.fragment))


def fetch_json(url: str, timeout: int = 10) -> dict[str, object]:
    request = Request(
        with_cache_buster(url),
        headers={
            "User-Agent": "WinSys-Guardian-Agent",
            "Accept": "application/json, application/vnd.github.raw+json",
            "Cache-Control": "no-cache",
        },
    )
    with urlopen(request, timeout=timeout) as response:
        data = json.loads(response.read().decode("utf-8"))
    return data if isinstance(data, dict) else {}


def download_file(url: str, destination: Path, timeout: int = 60) -> Path:
    request = Request(url, headers={"User-Agent": "WinSys-Guardian-Agent"})
    with urlopen(request, timeout=timeout) as response:
        with destination.open("wb") as handle:
            while True:
                chunk = response.read(1024 * 256)
                if not chunk:
                    break
                handle.write(chunk)
    return destination


def extract_agent_update(package_path: Path, destination: Path) -> tuple[Path | None, Path | None, Path | None]:
    destination.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(package_path, "r") as archive:
        archive.extractall(destination)

    candidates = list(destination.rglob("wga_agent.exe"))
    script_candidates = list(destination.rglob("wga_agent.py"))
    version_candidates = list(destination.rglob("version.json"))
    return (
        candidates[0] if candidates else None,
        script_candidates[0] if script_candidates else None,
        version_candidates[0] if version_candidates else None,
    )


def create_agent_update_helper(
    new_agent_exe: Path | None,
    new_agent_script: Path | None,
    new_version: Path | None,
) -> Path:
    helper_path = Path(tempfile.gettempdir()) / "wga_agent_update.cmd"
    lines = [
        "@echo off",
        "chcp 65001 >nul",
        "timeout /t 3 /nobreak >nul",
    ]
    if new_agent_exe is not None:
        lines.append(f'copy /Y "{new_agent_exe}" "{INSTALLED_AGENT_PATH}" >nul')
    if new_agent_script is not None:
        lines.append(f'copy /Y "{new_agent_script}" "{INSTALLED_SCRIPT_PATH}" >nul')
    if new_version is not None:
        lines.append(f'copy /Y "{new_version}" "{INSTALLED_VERSION_PATH}" >nul')
    if new_agent_exe is not None or INSTALLED_AGENT_PATH.exists():
        restart_line = f'start "" "{INSTALLED_AGENT_PATH}" --status --output "{DEFAULT_STATUS_PATH}"'
    else:
        restart_line = f'start "" "{sys.executable}" "{INSTALLED_SCRIPT_PATH}" --status --output "{DEFAULT_STATUS_PATH}"'
    lines.extend(
        [
            restart_line,
            "exit /b 0",
        ]
    )
    helper_path.write_text("\n".join(lines), encoding="utf-8")
    return helper_path


def check_and_install_agent_update() -> bool:
    info = local_version_info()
    current_version = str(info.get("version", "0.0.0")).strip() or "0.0.0"
    version_url = str(info.get("version_info_url", DEFAULT_VERSION_INFO_URL)).strip() or DEFAULT_VERSION_INFO_URL
    remote_info = fetch_json(version_url)
    latest_version = str(remote_info.get("version", "")).strip()
    package_url = str(remote_info.get("package_url", "")).strip()
    if not latest_version or not package_url:
        log_event("Update check skipped: remote version or package URL missing.")
        return False
    if normalize_version(latest_version) <= normalize_version(current_version):
        log_event(f"Agent is up to date: {current_version}.")
        return False

    work_dir = Path(tempfile.mkdtemp(prefix="wga_agent_update_"))
    package_path = download_file(package_url, work_dir / "WGA-portable.zip", timeout=120)
    new_agent_exe, new_agent_script, new_version = extract_agent_update(package_path, work_dir / "extracted")
    if new_agent_exe is None and new_agent_script is None:
        raise RuntimeError("Update пакетът не съдържа wga_agent.")
    helper_path = create_agent_update_helper(new_agent_exe, new_agent_script, new_version)
    log_event(f"Installing agent update {current_version} -> {latest_version}.")
    subprocess.Popen(
        ["cmd", "/c", str(helper_path)],
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    return True


def show_startup_splash(duration_ms: int = 3500) -> None:
    try:
        import tkinter as tk

        root = tk.Tk()
        root.overrideredirect(True)
        root.attributes("-topmost", True)
        root.configure(bg="#071311")
        width, height = 520, 260
        x = max(0, int((root.winfo_screenwidth() - width) / 2))
        y = max(0, int((root.winfo_screenheight() - height) / 2))
        root.geometry(f"{width}x{height}+{x}+{y}")

        frame = tk.Frame(root, bg="#071311", highlightbackground="#37e39a", highlightthickness=1)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        icon_path = bundled_file("assets/wga-icon.png")
        icon = None
        if icon_path.exists():
            try:
                icon = tk.PhotoImage(file=str(icon_path))
                icon = icon.subsample(max(1, int(max(icon.width(), icon.height()) / 82)))
                tk.Label(frame, image=icon, bg="#071311").pack(pady=(26, 10))
            except tk.TclError:
                icon = None
        if icon is None:
            tk.Label(frame, text="WGA", bg="#071311", fg="#37e39a", font=("Segoe UI Semibold", 30)).pack(pady=(32, 8))
        tk.Label(
            frame,
            text="WinSys Guardian Advanced",
            bg="#071311",
            fg="#ecfff7",
            font=("Segoe UI Semibold", 18),
        ).pack()
        tk.Label(
            frame,
            text="Компютърът е защитен и се наблюдава от WGA Agent",
            bg="#071311",
            fg="#a6d5c5",
            font=("Segoe UI", 10),
        ).pack(pady=(8, 0))
        root.after(duration_ms, root.destroy)
        root.mainloop()
    except Exception as exc:
        log_event(f"Startup splash failed: {exc}")


def validate_computer_name(name: str) -> tuple[bool, str]:
    value = name.strip().upper()
    if not value:
        return False, "Името не може да бъде празно."
    if len(value) > 15:
        return False, "Windows computer name може да бъде максимум 15 символа."
    allowed = set("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-")
    if any(char not in allowed for char in value):
        return False, "Позволени са само латински букви, цифри и тире."
    if value.startswith("-") or value.endswith("-"):
        return False, "Името не може да започва или завършва с тире."
    if value.isdigit():
        return False, "Името не може да съдържа само цифри."
    return True, value


def prompt_computer_name(current_name: str) -> str | None:
    import tkinter as tk
    from tkinter import messagebox, simpledialog

    root = tk.Tk()
    root.withdraw()
    if not messagebox.askyesno(
        "WGA Agent",
        f"Текущо име на компютъра: {current_name}\n\nИскате ли да го смените по време на инсталацията?",
    ):
        root.destroy()
        return None

    while True:
        requested = simpledialog.askstring(
            "Ново име на компютъра",
            "Въведете ново име до 15 символа\n(латински букви, цифри и тире):",
            initialvalue=current_name,
        )
        if requested is None:
            root.destroy()
            return None
        valid, result = validate_computer_name(requested)
        if valid:
            root.destroy()
            return result
        messagebox.showerror("Невалидно име", result)


def rename_computer_if_needed(new_name: str | None) -> tuple[bool, str]:
    if not new_name:
        return False, "Името на компютъра не е променено."
    current_name = socket.gethostname().upper()
    if new_name.upper() == current_name:
        return False, "Новото име съвпада с текущото."

    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        f"Rename-Computer -NewName '{new_name}' -Force",
    ]
    result = run_hidden(command, timeout=120)
    output = f"{result.stdout}\n{result.stderr}".strip()
    if result.returncode != 0:
        raise RuntimeError(output or "Rename-Computer не успя.")
    return True, "Името е променено. Нужен е рестарт, за да влезе в сила."


def installed_agent_command(*args: str) -> str:
    if getattr(sys, "frozen", False):
        agent_path = INSTALLED_AGENT_PATH
        return " ".join([f'"{agent_path}"', *args])
    agent_path = INSTALLED_SCRIPT_PATH
    return " ".join([f'"{sys.executable}"', f'"{agent_path}"', *args])


def relaunch_as_admin(*args: str) -> bool:
    if platform.system() != "Windows":
        return False
    if getattr(sys, "frozen", False):
        executable = str(INSTALLED_AGENT_PATH if INSTALLED_AGENT_PATH.exists() else Path(sys.executable))
        parameters = " ".join(args)
    else:
        script_path = INSTALLED_SCRIPT_PATH if INSTALLED_SCRIPT_PATH.exists() else Path(__file__).resolve()
        executable = sys.executable
        parameters = " ".join([f'"{script_path}"', *args])
    try:
        result = ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            executable,
            parameters,
            str(INSTALL_DIR if INSTALL_DIR.exists() else Path.cwd()),
            1,
        )
        return result > 32
    except OSError:
        return False


def copy_agent_to_install_dir() -> Path:
    INSTALL_DIR.mkdir(parents=True, exist_ok=True)
    REQUEST_DIR.mkdir(parents=True, exist_ok=True)
    source = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve()
    destination = INSTALLED_AGENT_PATH if getattr(sys, "frozen", False) else INSTALLED_SCRIPT_PATH
    if source != destination:
        shutil.copy2(source, destination)
    bundled_version = bundled_file("version.json")
    project_version = Path(__file__).resolve().parent / "version.json"
    if bundled_version.exists():
        shutil.copy2(bundled_version, INSTALLED_VERSION_PATH)
    elif project_version.exists():
        shutil.copy2(project_version, INSTALLED_VERSION_PATH)

    for icon_name in ("wga-icon.png", "wga-icon.ico"):
        bundled_icon = bundled_file(f"assets/{icon_name}")
        project_icon = Path(__file__).resolve().parent / "assets" / icon_name
        icon_target = INSTALL_DIR / "assets" / icon_name
        icon_target.parent.mkdir(parents=True, exist_ok=True)
        if bundled_icon.exists():
            shutil.copy2(bundled_icon, icon_target)
        elif project_icon.exists():
            shutil.copy2(project_icon, icon_target)
    return destination


def create_admin_console_shortcut() -> None:
    if platform.system() != "Windows":
        return
    ADMIN_SHORTCUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    icon_path = INSTALL_DIR / "assets" / "wga-icon.ico"
    if getattr(sys, "frozen", False) or INSTALLED_AGENT_PATH.exists():
        target_path = str(INSTALLED_AGENT_PATH)
        arguments = "--admin-console"
    else:
        target_path = sys.executable
        arguments = f'"{INSTALLED_SCRIPT_PATH}" --admin-console'

    script = f"""
$shortcutPath = {json.dumps(str(ADMIN_SHORTCUT_PATH))}
$targetPath = {json.dumps(target_path)}
$arguments = {json.dumps(arguments)}
$workingDirectory = {json.dumps(str(INSTALL_DIR))}
$iconPath = {json.dumps(str(icon_path if icon_path.exists() else target_path))}
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $targetPath
$shortcut.Arguments = $arguments
$shortcut.WorkingDirectory = $workingDirectory
$shortcut.Hotkey = "CTRL+ALT+W"
$shortcut.Description = "WGA Agent Admin Console"
$shortcut.IconLocation = $iconPath
$shortcut.Save()
icacls $shortcutPath /inheritance:r /grant "*S-1-5-32-544:F" /grant "SYSTEM:F" | Out-Null
"""
    result = run_hidden(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            script,
        ],
        timeout=60,
    )
    if result.returncode != 0:
        output = f"{result.stdout}\n{result.stderr}".strip()
        log_event(f"Failed to create admin shortcut: {output}")


def remove_admin_console_shortcut() -> None:
    try:
        if ADMIN_SHORTCUT_PATH.exists():
            ADMIN_SHORTCUT_PATH.unlink()
    except OSError as exc:
        log_event(f"Failed to remove admin shortcut: {exc}")


def configure_request_permissions() -> None:
    if platform.system() != "Windows":
        return
    result = run_hidden(
        [
            "icacls",
            str(REQUEST_DIR),
            "/grant",
            "*S-1-5-32-545:(OI)(CI)M",
            "/T",
            "/C",
        ],
        timeout=60,
    )
    if result.returncode != 0:
        output = f"{result.stdout}\n{result.stderr}".strip()
        log_event(f"Failed to configure request permissions: {output}")


def register_status_task(status_path: Path, interval_minutes: int) -> None:
    task_command = installed_agent_command("--maintenance", "--output", f'"{status_path}"')
    result = run_hidden(
        [
            "schtasks",
            "/Create",
            "/TN",
            STATUS_TASK_NAME,
            "/SC",
            "MINUTE",
            "/MO",
            str(interval_minutes),
            "/TR",
            task_command,
            "/RU",
            "SYSTEM",
            "/RL",
            "HIGHEST",
            "/F",
        ],
        timeout=60,
    )
    output = f"{result.stdout}\n{result.stderr}".strip()
    if result.returncode != 0:
        raise RuntimeError(output or "Scheduled task регистрацията не успя.")


def register_startup_task() -> None:
    task_command = installed_agent_command("--startup")
    result = run_hidden(
        [
            "schtasks",
            "/Create",
            "/TN",
            STARTUP_TASK_NAME,
            "/SC",
            "ONLOGON",
            "/TR",
            task_command,
            "/RL",
            "HIGHEST",
            "/F",
        ],
        timeout=60,
    )
    output = f"{result.stdout}\n{result.stderr}".strip()
    if result.returncode != 0:
        raise RuntimeError(output or "Startup task регистрацията не успя.")


def unregister_status_task(remove_shortcut: bool = False) -> None:
    run_hidden(["schtasks", "/Delete", "/TN", STATUS_TASK_NAME, "/F"], timeout=60)
    run_hidden(["schtasks", "/Delete", "/TN", STARTUP_TASK_NAME, "/F"], timeout=60)
    if remove_shortcut:
        remove_admin_console_shortcut()


def show_install_result(message: str) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("WGA Agent", message)
        root.destroy()
    except Exception:
        print(message)


def read_text_tail(path: Path, max_chars: int = 12000) -> str:
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return ""
    return text[-max_chars:]


def query_task_status(task_name: str) -> str:
    if platform.system() != "Windows":
        return "Windows only"
    result = run_hidden(["schtasks", "/Query", "/TN", task_name, "/FO", "LIST", "/V"], timeout=60)
    output = f"{result.stdout}\n{result.stderr}".strip()
    return output or "No task information."


def format_status_summary(data: dict[str, object]) -> str:
    hardware = data.get("hardware_info", {}) if isinstance(data.get("hardware_info"), dict) else {}
    system_info = hardware.get("computer_system", {}) if isinstance(hardware.get("computer_system"), dict) else {}
    memory = hardware.get("memory", {}) if isinstance(hardware.get("memory"), dict) else {}
    update_status = data.get("update_status", {}) if isinstance(data.get("update_status"), dict) else {}
    lines = [
        f"Computer: {data.get('hostname', 'N/A')}",
        f"Timestamp: {data.get('timestamp', 'N/A')}",
        f"Platform: {data.get('platform', 'N/A')} {data.get('platform_release', '')} {data.get('platform_version', '')}",
        f"Model: {system_info.get('manufacturer', 'N/A')} {system_info.get('model', '')}",
        f"Online: {'Yes' if data.get('online') else 'No'}",
        f"Local IPs: {', '.join(data.get('local_ips', [])) or 'N/A'}",
        f"Uptime: {data.get('system_uptime', 'N/A')}",
        f"RAM: {memory.get('used_gb', 'N/A')} / {memory.get('total_gb', 'N/A')} GB",
        f"Pending Windows Updates: {update_status.get('pending_updates', 'N/A')}",
        f"Status file: {DEFAULT_STATUS_PATH}",
        f"Request queue: {COMMAND_QUEUE_PATH}",
    ]
    return "\n".join(lines)


def format_live_activity(data: dict[str, object]) -> str:
    lines = [
        f"Timestamp: {data.get('timestamp', 'N/A')}",
        f"Interactive user: {data.get('interactive_user', 'N/A') or 'N/A'}",
        "",
        "== Foreground window ==",
    ]
    foreground = data.get("foreground_window", {})
    if isinstance(foreground, dict) and foreground:
        lines.extend(
            [
                f"Process: {foreground.get('process_name', 'N/A')} ({foreground.get('process_id', 'N/A')})",
                f"Title: {foreground.get('window_title', '') or 'N/A'}",
            ]
        )
    else:
        lines.append("N/A")

    lines.extend(["", "== Top processes =="])
    for process in data.get("top_processes", []) if isinstance(data.get("top_processes"), list) else []:
        if not isinstance(process, dict):
            continue
        title = str(process.get("window_title", "") or "")
        title_suffix = f" | {title}" if title else ""
        lines.append(
            f"{process.get('name', 'N/A')} pid={process.get('pid', 'N/A')} "
            f"cpu={process.get('cpu_seconds', 'N/A')}s mem={process.get('memory_mb', 'N/A')}MB{title_suffix}"
        )

    lines.extend(["", "== Not responding processes =="])
    hung_processes = data.get("hung_processes", []) if isinstance(data.get("hung_processes"), list) else []
    if not hung_processes:
        lines.append("None detected.")
    for process in hung_processes:
        if not isinstance(process, dict):
            continue
        title = str(process.get("window_title", "") or "")
        title_suffix = f" | {title}" if title else ""
        lines.append(f"{process.get('name', 'N/A')} pid={process.get('pid', 'N/A')}{title_suffix}")

    lines.extend(["", "== Established network connections =="])
    for connection in data.get("network_connections", []) if isinstance(data.get("network_connections"), list) else []:
        if not isinstance(connection, dict):
            continue
        lines.append(
            f"{connection.get('process_name', 'N/A')} pid={connection.get('process_id', 'N/A')} "
            f"{connection.get('local', 'N/A')} -> {connection.get('remote', 'N/A')}"
        )

    lines.extend(["", "== DNS cache domains =="])
    for entry in data.get("dns_cache", []) if isinstance(data.get("dns_cache"), list) else []:
        if not isinstance(entry, dict):
            continue
        lines.append(f"{entry.get('domain', 'N/A')} [{entry.get('type', 'N/A')}] {entry.get('data', '')}")

    lines.extend(["", "== Remote Desktop =="])
    rdp = data.get("remote_desktop", {}) if isinstance(data.get("remote_desktop"), dict) else {}
    if not rdp:
        lines.append("N/A")
    else:
        lines.extend(
            [
                f"Enabled: {'Yes' if rdp.get('enabled') else 'No'}",
                f"NLA: {'Yes' if rdp.get('nla_enabled') else 'No'}",
                f"Firewall rules enabled: {rdp.get('firewall_rules_enabled', 'N/A')}",
                f"Host: {rdp.get('hostname', 'N/A')}",
                f"IPs: {', '.join(rdp.get('local_ips', [])) if isinstance(rdp.get('local_ips'), list) else 'N/A'}",
                f"Connect: {rdp.get('connect_command', 'N/A')}",
            ]
        )
        if rdp.get("connect_command_ip"):
            lines.append(f"Connect by IP: {rdp.get('connect_command_ip')}")
        lines.append("")
        lines.append("Sessions:")
        sessions = rdp.get("sessions", []) if isinstance(rdp.get("sessions"), list) else []
        if not sessions:
            lines.append("No sessions detected.")
        for session in sessions:
            if not isinstance(session, dict):
                continue
            lines.append(
                f"id={session.get('id', 'N/A')} state={session.get('state', 'N/A')} "
                f"user={session.get('username', '') or 'N/A'} name={session.get('session_name', 'N/A')}"
            )
        shadow_commands = rdp.get("shadow_commands", []) if isinstance(rdp.get("shadow_commands"), list) else []
        if shadow_commands:
            lines.append("")
            lines.append("Shadow commands:")
            lines.extend(str(command) for command in shadow_commands)

    return "\n".join(lines)


def show_admin_console() -> None:
    if platform.system() != "Windows":
        show_install_result("WGA Agent Admin Console е предназначена за Windows.")
        return
    if not is_windows_admin():
        if not relaunch_as_admin("--admin-console"):
            show_install_result("Admin Console изисква Administrator права.")
        return

    import tkinter as tk
    from tkinter import messagebox, simpledialog, ttk

    root = tk.Tk()
    root.title("WGA Agent Admin Console")
    root.geometry("980x720")
    root.minsize(820, 560)
    root.configure(bg="#071311")

    icon_path = INSTALL_DIR / "assets" / "wga-icon.ico"
    if icon_path.exists():
        try:
            root.iconbitmap(str(icon_path))
        except tk.TclError:
            pass

    header = tk.Frame(root, bg="#0d1c1a", height=78, highlightbackground="#1f4e46", highlightthickness=1)
    header.pack(fill="x")
    header.pack_propagate(False)
    tk.Label(
        header,
        text="WGA Agent Admin Console",
        bg="#0d1c1a",
        fg="#ecfff7",
        font=("Segoe UI Semibold", 20),
    ).pack(anchor="w", padx=22, pady=(13, 0))
    tk.Label(
        header,
        text="Administrator-only status, maintenance and stop controls",
        bg="#0d1c1a",
        fg="#a6d5c5",
        font=("Segoe UI", 10),
    ).pack(anchor="w", padx=23, pady=(2, 0))

    body = tk.Frame(root, bg="#071311")
    body.pack(fill="both", expand=True, padx=18, pady=16)

    toolbar = tk.Frame(body, bg="#071311")
    toolbar.pack(fill="x", pady=(0, 12))

    notebook = ttk.Notebook(body)
    notebook.pack(fill="both", expand=True)

    tabs: dict[str, tk.Text] = {}
    for tab_name in ("Summary", "Live Activity", "Full Status JSON", "Tasks", "Command Results", "Log"):
        frame = tk.Frame(notebook, bg="#071311")
        text = tk.Text(
            frame,
            bg="#07100f",
            fg="#d8fff1",
            insertbackground="#d8fff1",
            relief="flat",
            wrap="word",
            font=("Consolas", 10),
            padx=12,
            pady=10,
        )
        scroll = tk.Scrollbar(frame, command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        notebook.add(frame, text=tab_name)
        tabs[tab_name] = text

    def set_tab(name: str, content: str) -> None:
        text = tabs[name]
        text.configure(state="normal")
        text.delete("1.0", "end")
        text.insert("end", content)
        text.configure(state="disabled")

    def refresh() -> None:
        status_data = read_json_file(DEFAULT_STATUS_PATH)
        if not status_data:
            try:
                save_status(DEFAULT_STATUS_PATH)
                status_data = read_json_file(DEFAULT_STATUS_PATH)
            except Exception as exc:
                status_data = {"error": str(exc)}

        set_tab("Summary", format_status_summary(status_data))
        try:
            activity_data = collect_realtime_activity()
            REALTIME_ACTIVITY_PATH.write_text(json.dumps(activity_data, indent=2), encoding="utf-8")
        except Exception as exc:
            activity_data = read_json_file(REALTIME_ACTIVITY_PATH) or {"error": str(exc)}
        set_tab("Live Activity", format_live_activity(activity_data))
        set_tab("Full Status JSON", json.dumps(status_data, indent=2, ensure_ascii=False))
        set_tab(
            "Tasks",
            "\n\n".join(
                [
                    f"== {STATUS_TASK_NAME} ==",
                    query_task_status(STATUS_TASK_NAME),
                    f"== {STARTUP_TASK_NAME} ==",
                    query_task_status(STARTUP_TASK_NAME),
                ]
            ),
        )
        set_tab(
            "Command Results",
            json.dumps(read_json_file(COMMAND_RESULT_PATH) or {"results": []}, indent=2, ensure_ascii=False),
        )
        set_tab("Log", read_text_tail(LOG_PATH) or "No agent log yet.")

    def run_maintenance_now() -> None:
        try:
            run_maintenance_sequence(DEFAULT_STATUS_PATH)
            refresh()
            messagebox.showinfo("WGA Agent", "Maintenance изпълнена успешно.", parent=root)
        except Exception as exc:
            messagebox.showerror("WGA Agent", f"Maintenance грешка:\n{exc}", parent=root)

    def request_kill_hung_process() -> None:
        value = simpledialog.askstring(
            "Kill Hung Process",
            "Въведете PID на забил процес от Live Activity:",
            parent=root,
        )
        if value is None:
            return
        if not value.strip().isdigit():
            messagebox.showerror("Kill Hung Process", "PID трябва да бъде число.", parent=root)
            return
        queue = read_json_file(COMMAND_QUEUE_PATH)
        commands = queue.get("commands", [])
        if not isinstance(commands, list):
            commands = []
        command_id = str(uuid.uuid4())
        commands.append({"id": command_id, "type": "kill_hung_process", "pid": int(value.strip())})
        COMMAND_QUEUE_PATH.parent.mkdir(parents=True, exist_ok=True)
        COMMAND_QUEUE_PATH.write_text(json.dumps({"commands": commands}, indent=2), encoding="utf-8")
        run_maintenance_sequence(DEFAULT_STATUS_PATH)
        refresh()
        messagebox.showinfo("Kill Hung Process", f"Заявката е обработена: {command_id}", parent=root)

    def request_enable_rdp() -> None:
        if not messagebox.askyesno(
            "Enable Remote Desktop",
            "Това ще включи Windows Remote Desktop, NLA и firewall правилата на машината. Продължаване?",
            parent=root,
        ):
            return
        queue = read_json_file(COMMAND_QUEUE_PATH)
        commands = queue.get("commands", [])
        if not isinstance(commands, list):
            commands = []
        command_id = str(uuid.uuid4())
        commands.append({"id": command_id, "type": "enable_remote_desktop"})
        COMMAND_QUEUE_PATH.parent.mkdir(parents=True, exist_ok=True)
        COMMAND_QUEUE_PATH.write_text(json.dumps({"commands": commands}, indent=2), encoding="utf-8")
        run_maintenance_sequence(DEFAULT_STATUS_PATH)
        refresh()
        messagebox.showinfo("Enable Remote Desktop", f"Заявката е обработена: {command_id}", parent=root)

    auto_refresh_var = tk.BooleanVar(value=False)
    auto_refresh_job: str | None = None

    def schedule_auto_refresh() -> None:
        nonlocal auto_refresh_job
        if auto_refresh_var.get():
            refresh()
            auto_refresh_job = root.after(5000, schedule_auto_refresh)
        else:
            auto_refresh_job = None

    def toggle_auto_refresh() -> None:
        nonlocal auto_refresh_job
        if auto_refresh_job is not None:
            root.after_cancel(auto_refresh_job)
            auto_refresh_job = None
        schedule_auto_refresh()

    def stop_agent() -> None:
        if not messagebox.askyesno(
            "Stop WGA Agent",
            "Това ще спре WGA Agent, като премахне scheduled tasks. Продължаване?",
            parent=root,
        ):
            return
        unregister_status_task()
        refresh()
        messagebox.showinfo("WGA Agent", "Agent scheduled tasks са премахнати.", parent=root)

    tk.Button(
        toolbar,
        text="Refresh",
        command=refresh,
        bg="#173c36",
        activebackground="#1f554c",
        fg="#d9fff1",
        relief="flat",
        padx=16,
        pady=8,
        font=("Segoe UI Semibold", 10),
    ).pack(side="left")
    tk.Button(
        toolbar,
        text="Run Maintenance",
        command=run_maintenance_now,
        bg="#1d527f",
        activebackground="#2f8fff",
        fg="#eef7ff",
        relief="flat",
        padx=16,
        pady=8,
        font=("Segoe UI Semibold", 10),
    ).pack(side="left", padx=(8, 0))
    tk.Button(
        toolbar,
        text="Kill Hung Process",
        command=request_kill_hung_process,
        bg="#6f5a1d",
        activebackground="#a47e32",
        fg="#fff8e5",
        relief="flat",
        padx=16,
        pady=8,
        font=("Segoe UI Semibold", 10),
    ).pack(side="left", padx=(8, 0))
    tk.Button(
        toolbar,
        text="Enable RDP",
        command=request_enable_rdp,
        bg="#245634",
        activebackground="#2f7044",
        fg="#f3fff5",
        relief="flat",
        padx=16,
        pady=8,
        font=("Segoe UI Semibold", 10),
    ).pack(side="left", padx=(8, 0))
    tk.Checkbutton(
        toolbar,
        text="Auto refresh 5s",
        variable=auto_refresh_var,
        command=toggle_auto_refresh,
        bg="#071311",
        activebackground="#071311",
        fg="#a6d5c5",
        activeforeground="#ecfff7",
        selectcolor="#0d1c1a",
        font=("Segoe UI", 10),
    ).pack(side="left", padx=(12, 0))
    tk.Button(
        toolbar,
        text="Stop Agent",
        command=stop_agent,
        bg="#7f2f39",
        activebackground="#c94d58",
        fg="#fff5f5",
        relief="flat",
        padx=16,
        pady=8,
        font=("Segoe UI Semibold", 10),
    ).pack(side="right")

    refresh()
    root.mainloop()


def install_agent(computer_name: str | None = None, interval_minutes: int = DEFAULT_INTERVAL_MINUTES) -> Path:
    if platform.system() != "Windows":
        raise RuntimeError("WGA Agent installation е предназначена за Windows.")
    if not is_windows_admin():
        raise RuntimeError("Стартирайте инсталацията като Administrator.")

    requested_name = computer_name
    if requested_name is None:
        requested_name = prompt_computer_name(socket.gethostname())
    elif requested_name:
        valid, normalized_name = validate_computer_name(requested_name)
        if not valid:
            raise RuntimeError(normalized_name)
        requested_name = normalized_name

    rename_changed, rename_message = rename_computer_if_needed(requested_name)
    copy_agent_to_install_dir()
    configure_request_permissions()
    register_status_task(DEFAULT_STATUS_PATH, interval_minutes)
    register_startup_task()
    create_admin_console_shortcut()
    save_status(DEFAULT_STATUS_PATH)

    installed_message = (
        f"{APP_NAME} е инсталиран успешно.\n\n"
        f"Status файл: {DEFAULT_STATUS_PATH}\n"
        f"Scheduled tasks: {STATUS_TASK_NAME}, {STARTUP_TASK_NAME}\n"
        "Admin shortcut: Ctrl+Alt+W\n"
        f"Интервал: {interval_minutes} мин."
    )
    if rename_changed:
        installed_message += f"\n\n{rename_message}"
    show_install_result(installed_message)
    return DEFAULT_STATUS_PATH


def run_startup_sequence() -> None:
    show_startup_splash()


def run_maintenance_sequence(output_path: Path | str = DEFAULT_STATUS_PATH) -> None:
    try:
        if check_and_install_agent_update():
            return
    except Exception as exc:
        log_event(f"Agent update failed: {exc}")
    try:
        process_command_queue()
    except Exception as exc:
        log_event(f"Command queue processing failed: {exc}")
    try:
        save_status(output_path)
    except Exception as exc:
        log_event(f"Maintenance status update failed: {exc}")
    try:
        status_path = Path(output_path)
        realtime_path = REALTIME_ACTIVITY_PATH if status_path == DEFAULT_STATUS_PATH else status_path.with_name("wga_realtime_activity.json")
        save_realtime_activity(realtime_path)
    except Exception as exc:
        log_event(f"Realtime activity update failed: {exc}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Collect Windows agent status and save it to a JSON file."
    )
    parser.add_argument(
        "--install",
        action="store_true",
        help="Install the WGA monitoring agent and register its scheduled task.",
    )
    parser.add_argument(
        "--uninstall",
        action="store_true",
        help="Remove the WGA monitoring agent scheduled task.",
    )
    parser.add_argument(
        "--status",
        action="store_true",
        help="Collect one status snapshot and exit.",
    )
    parser.add_argument(
        "--realtime",
        action="store_true",
        help="Collect one realtime activity snapshot and exit.",
    )
    parser.add_argument(
        "--startup",
        action="store_true",
        help="Show the startup splash for the interactive user.",
    )
    parser.add_argument(
        "--maintenance",
        action="store_true",
        help="Run elevated maintenance: update check and one status snapshot.",
    )
    parser.add_argument(
        "--check-update",
        action="store_true",
        help="Check for an agent update and install it when available.",
    )
    parser.add_argument(
        "--admin-console",
        action="store_true",
        help="Open the administrator-only WGA Agent console.",
    )
    parser.add_argument(
        "--computer-name",
        default=None,
        help="Optional new Windows computer name to apply during installation.",
    )
    parser.add_argument(
        "--interval-minutes",
        type=int,
        default=DEFAULT_INTERVAL_MINUTES,
        help="Scheduled status refresh interval in minutes.",
    )
    parser.add_argument(
        "--output",
        default=str(DEFAULT_STATUS_PATH),
        help="Path to output status JSON file.",
    )
    args = parser.parse_args()

    try:
        if args.uninstall:
            unregister_status_task(remove_shortcut=True)
            show_install_result(f"{APP_NAME} scheduled task е премахнат.")
        elif args.install:
            output_file = install_agent(args.computer_name, max(1, args.interval_minutes))
            print(f"Agent installed. Status file: {output_file}")
        elif args.startup:
            run_startup_sequence()
        elif args.maintenance:
            run_maintenance_sequence(args.output)
        elif args.realtime:
            output_file = save_realtime_activity(args.output)
            print(f"Realtime activity saved to: {output_file}")
        elif args.check_update:
            updated = check_and_install_agent_update()
            print("Agent update started." if updated else "Agent is up to date.")
        elif args.admin_console:
            show_admin_console()
        else:
            output_file = save_status(args.output)
            print(f"Agent status saved to: {output_file}")
    except Exception as exc:
        show_install_result(f"Грешка: {exc}")
        raise
