from __future__ import annotations

import json

from services.paths import runtime_file


def current_version() -> str:
    try:
        data = json.loads(runtime_file("version.json").read_text(encoding="utf-8-sig"))
    except Exception:
        return "unknown"
    return str(data.get("version") or "unknown")
