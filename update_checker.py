from __future__ import annotations

import json
from dataclasses import dataclass
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlsplit, urlunsplit
from urllib.request import urlopen


@dataclass
class UpdateResult:
    status: str
    latest_version: str = ""
    download_url: str = ""
    package_url: str = ""
    changelog: tuple[str, ...] = ()
    notes: str = ""
    error: str = ""


def _normalize_version(version: str) -> tuple[int, ...]:
    parts: list[int] = []
    for item in version.strip().split("."):
        try:
            parts.append(int(item))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def _fetch_json(url: str, timeout: int = 6) -> dict[str, str]:
    prepared_url = _prepare_url(url)
    with urlopen(prepared_url, timeout=timeout) as response:
        raw = response.read().decode("utf-8")
    data = json.loads(raw)
    return data if isinstance(data, dict) else {}


def _prepare_url(url: str) -> str:
    stripped_url = url.strip()
    if not stripped_url:
        raise ValueError("Missing update URL.")

    parts = urlsplit(stripped_url)
    if parts.scheme not in {"http", "https"} or not parts.netloc:
        raise ValueError("Invalid update URL. Use a full http/https GitHub raw link.")

    encoded_path = quote(parts.path, safe="/-._~")
    encoded_query = quote(parts.query, safe="=&-._~")
    encoded_fragment = quote(parts.fragment, safe="-._~")
    return urlunsplit((parts.scheme, parts.netloc, encoded_path, encoded_query, encoded_fragment))


def check_for_updates(current_version: str, version_info_url: str) -> UpdateResult:
    if not version_info_url.strip():
        return UpdateResult(status="not_configured")

    try:
        remote_info = _fetch_json(version_info_url.strip())
    except HTTPError as exc:
        if exc.code == 404:
            return UpdateResult(
                status="raw_unavailable",
                error="GitHub raw version.json is not publicly available yet.",
            )
        return UpdateResult(status="error", error=str(exc))
    except (URLError, json.JSONDecodeError, TimeoutError, ValueError) as exc:
        return UpdateResult(status="error", error=str(exc))

    latest_version = str(remote_info.get("version", "")).strip()
    download_url = str(remote_info.get("download_url", "")).strip()
    package_url = str(remote_info.get("package_url", "")).strip()
    notes = str(remote_info.get("notes", "")).strip()
    raw_changelog = remote_info.get("changelog", [])
    changelog = tuple(str(item).strip() for item in raw_changelog if str(item).strip()) if isinstance(raw_changelog, list) else ()

    if not latest_version:
        return UpdateResult(status="error", error="Remote version metadata is missing a version field.")

    if _normalize_version(latest_version) > _normalize_version(current_version):
        return UpdateResult(
            status="update_available",
            latest_version=latest_version,
            download_url=download_url,
            package_url=package_url,
            changelog=changelog,
            notes=notes,
        )

    return UpdateResult(
        status="up_to_date",
        latest_version=latest_version,
        download_url=download_url,
        package_url=package_url,
        changelog=changelog,
        notes=notes,
    )
