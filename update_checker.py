# Проверява online version metadata и сравнява версиите.
from __future__ import annotations

import json
import time
from dataclasses import dataclass
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qsl, quote, urlencode, urlsplit, urlunsplit
from urllib.request import Request, urlopen


# Описва данните, които приложението пази за UpdateResult.
@dataclass
class UpdateResult:
    # Пази резултата от online проверката за нова версия.
    status: str
    latest_version: str = ""
    download_url: str = ""
    package_url: str = ""
    changelog: tuple[str, ...] = ()
    notes: str = ""
    error: str = ""


# Помощна функция за normalize version.
def _normalize_version(version: str) -> tuple[int, ...]:
    # Превръща версия като 0.2.4 в числа за сравнение.
    parts: list[int] = []
    for item in version.strip().split("."):
        try:
            parts.append(int(item))
        except ValueError:
            parts.append(0)
    return tuple(parts)


# Помощна функция за fetch json.
def _fetch_json(url: str, timeout: int = 6) -> dict[str, str]:
    # Изтегля JSON файла с информация за последната версия.
    prepared_url = _prepare_url(url)
    prepared_url = _with_cache_buster(prepared_url)
    request = Request(
        prepared_url,
        headers={
            "User-Agent": "WinSys-Guardian-Advanced-Updater",
            "Accept": "application/json, application/vnd.github.raw+json",
            "Cache-Control": "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
        },
    )
    with urlopen(request, timeout=timeout) as response:
        raw = response.read().decode("utf-8")
    data = json.loads(raw)
    return data if isinstance(data, dict) else {}


def _candidate_urls(version_info_url: str) -> tuple[str, ...]:
    """Build independent GitHub endpoints so one stale CDN response cannot hide an update."""
    primary = version_info_url.strip()
    candidates = [primary]
    parts = urlsplit(primary)
    if parts.netloc.lower() == "raw.githubusercontent.com":
        path_parts = [item for item in parts.path.split("/") if item]
        if len(path_parts) >= 4:
            owner, repository = path_parts[0], path_parts[1]
            if path_parts[2:5] == ["refs", "heads", "main"]:
                file_parts = path_parts[5:]
            else:
                file_parts = path_parts[3:]
            file_path = "/".join(file_parts) or "version.json"
            candidates.extend(
                (
                    f"https://raw.githubusercontent.com/{owner}/{repository}/main/{file_path}",
                    f"https://api.github.com/repos/{owner}/{repository}/contents/{file_path}?ref=main",
                )
            )
    return tuple(dict.fromkeys(candidates))


def _fetch_latest_json(version_info_url: str) -> dict[str, object]:
    responses: list[dict[str, object]] = []
    errors: list[Exception] = []
    for candidate in _candidate_urls(version_info_url):
        try:
            data = _fetch_json(candidate)
            if str(data.get("version", "")).strip():
                responses.append(data)
        except (HTTPError, URLError, json.JSONDecodeError, TimeoutError, ValueError) as exc:
            errors.append(exc)
    if responses:
        return max(responses, key=lambda item: _normalize_version(str(item.get("version", ""))))
    if errors:
        raise errors[-1]
    return {}


# Помощна функция за prepare url.
def _prepare_url(url: str) -> str:
    # Проверява дали update адресът е валиден пълен URL.
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


# Помощна функция за with cache buster.
def _with_cache_buster(url: str) -> str:
    # Добавя време към адреса, за да заобиколим GitHub кеша.
    parts = urlsplit(url)
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    query["_wga_ts"] = str(int(time.time()))
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query), parts.fragment))


# Проверява дали има по-нова версия на приложението.
def check_for_updates(current_version: str, version_info_url: str) -> UpdateResult:
    # Сравнява локалната версия с online version.json файла.
    if not version_info_url.strip():
        return UpdateResult(status="not_configured")

    try:
        remote_info = _fetch_latest_json(version_info_url.strip())
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
