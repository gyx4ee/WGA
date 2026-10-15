from __future__ import annotations

import re

from legacy_app import SECURE_STORE_FILE, load_secure_store, normalize_product_key_input, save_secure_store


WINDOWS_KEYS = {
    "save_windows10_key": ("Windows 10", "windows10_product_key"),
    "show_windows10_key": ("Windows 10", "windows10_product_key"),
    "clear_windows10_key": ("Windows 10", "windows10_product_key"),
    "save_windows11_key": ("Windows 11", "windows11_product_key"),
    "show_windows11_key": ("Windows 11", "windows11_product_key"),
    "clear_windows11_key": ("Windows 11", "windows11_product_key"),
}

OFFICE_VERSION_CHOICES = {
    "Office 2016": "office_2016_activation",
    "Office 2019": "office_2019_activation",
    "Office 2021": "office_2021_activation",
}

OFFICE_PRODUCT_CHOICES = [
    "Office 2016",
    "Office 2019",
    "Office 2021",
    "Office 2024",
    "Microsoft 365",
    "Project 2019",
    "Project 2021",
    "Project 2024",
    "Visio 2019",
    "Visio 2021",
    "Visio 2024",
]

OFFICE_LEGACY_STORE_KEYS = {
    "Office 2016": "office_2016_activation_product_key",
    "Office 2019": "office_2019_activation_product_key",
    "Office 2021": "office_2021_activation_product_key",
}

WINDOWS_GENERIC_KEY_LABEL = "Общ ключ"
WINDOWS_EDITION_CHOICES = [
    WINDOWS_GENERIC_KEY_LABEL,
    "Home",
    "Home N",
    "Pro",
    "Pro N",
    "Pro for Workstations",
    "Enterprise",
    "Enterprise N",
    "Enterprise LTSC",
    "IoT Enterprise",
    "Education",
    "Education N",
    "SE",
]


def normalize_product_key(raw_value: str) -> str:
    return normalize_product_key_input(raw_value)


def read_store() -> dict[str, str]:
    return load_secure_store()


def save_key(store_key: str, product_key: str) -> None:
    store = read_store()
    store[store_key] = product_key
    save_secure_store(store)


def get_key(store_key: str) -> str:
    return read_store().get(store_key, "").strip()


def clear_key(store_key: str) -> bool:
    store = read_store()
    if store_key not in store:
        return False
    store.pop(store_key, None)
    save_secure_store(store)
    return True


def secure_store_path() -> str:
    return str(SECURE_STORE_FILE)


def windows_key_choices() -> list[str]:
    return list(WINDOWS_EDITION_CHOICES)


def windows_store_key(version_label: str, edition_label: str) -> str:
    version_slug = version_label.lower().replace(" ", "")
    if edition_label == WINDOWS_GENERIC_KEY_LABEL:
        return f"{version_slug}_product_key"
    edition_slug = re.sub(r"[^a-z0-9]+", "_", edition_label.lower()).strip("_")
    return f"{version_slug}_{edition_slug}_product_key"


def windows_key_label(version_label: str, edition_label: str) -> str:
    if edition_label == WINDOWS_GENERIC_KEY_LABEL:
        return f"{version_label} общ ключ"
    return f"{version_label} {edition_label}"


def office_key_choices() -> list[str]:
    return list(OFFICE_PRODUCT_CHOICES)


def office_store_key(product_label: str) -> str:
    product_slug = re.sub(r"[^a-z0-9]+", "_", product_label.lower()).strip("_")
    return f"{product_slug}_product_key"


def office_legacy_store_key(product_label: str) -> str | None:
    return OFFICE_LEGACY_STORE_KEYS.get(product_label)


def office_key_label(product_label: str) -> str:
    return product_label.strip()


def normalize_windows_edition(raw_edition: str) -> str:
    normalized = raw_edition.strip().lower().replace("-", " ")
    aliases = {
        "professional": "Pro",
        "professionaln": "Pro N",
        "pro": "Pro",
        "core": "Home",
        "home": "Home",
        "enterprise": "Enterprise",
        "enterprises": "Enterprise",
        "enterprisen": "Enterprise N",
        "enterprise ltsc": "Enterprise LTSC",
        "iotenterprise": "IoT Enterprise",
        "iot enterprise": "IoT Enterprise",
        "education": "Education",
        "educationn": "Education N",
        "cloud": "SE",
        "se": "SE",
    }
    compact = normalized.replace(" ", "")
    return aliases.get(normalized, aliases.get(compact, raw_edition.strip() or WINDOWS_GENERIC_KEY_LABEL))
