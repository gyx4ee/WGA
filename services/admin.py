from __future__ import annotations

import ctypes
import subprocess
import sys
from pathlib import Path


def is_running_as_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin() -> bool:
    if getattr(sys, "frozen", False):
        target = str(Path(sys.executable).resolve())
        args = subprocess.list2cmdline(sys.argv[1:])
    else:
        target = sys.executable
        args = subprocess.list2cmdline([str(Path(sys.argv[0]).resolve()), *sys.argv[1:]])
    result = ctypes.windll.shell32.ShellExecuteW(None, "runas", target, args, None, 1)
    return result > 32
