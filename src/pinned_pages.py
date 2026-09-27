"""Local pinned-page preferences for DataLens navigation."""

from __future__ import annotations

import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable


MAX_PINNED_PAGES = 10
PINNED_PAGES_ENV_VAR = "DATALENS_PINNED_PAGES_FILE"


def get_pinned_pages_path(path: str | os.PathLike[str] | None = None) -> Path:
    """Return the local file used to persist pinned pages."""
    if path is not None:
        return Path(path)

    env_path = os.environ.get(PINNED_PAGES_ENV_VAR)
    if env_path:
        return Path(env_path)

    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "DataLens" / "pinned_pages.json"

    return Path.home() / ".datalens" / "pinned_pages.json"


def normalize_pinned_pages(page_keys: Iterable[str], valid_page_keys: Iterable[str]) -> list[str]:
    """Keep pinned pages valid, unique, ordered, and within the pin limit."""
    valid = set(valid_page_keys)
    normalized: list[str] = []

    for page_key in page_keys:
        if page_key in valid and page_key not in normalized:
            normalized.append(page_key)
        if len(normalized) >= MAX_PINNED_PAGES:
            break

    return normalized


def load_pinned_pages(
    valid_page_keys: Iterable[str],
    path: str | os.PathLike[str] | None = None,
) -> list[str]:
    """Load pinned page IDs from local disk."""
    pinned_path = get_pinned_pages_path(path)
    if not pinned_path.exists():
        return []

    try:
        raw = json.loads(pinned_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []

    page_keys = raw.get("pinned_pages", []) if isinstance(raw, dict) else raw
    if not isinstance(page_keys, list):
        return []

    return normalize_pinned_pages(page_keys, valid_page_keys)


def save_pinned_pages(
    page_keys: Iterable[str],
    valid_page_keys: Iterable[str],
    path: str | os.PathLike[str] | None = None,
) -> list[str]:
    """Persist pinned page IDs to local disk and return the saved list."""
    pinned_pages = normalize_pinned_pages(page_keys, valid_page_keys)
    pinned_path = get_pinned_pages_path(path)
    pinned_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": 1,
        "updated_at": datetime.now(timezone.utc).isoformat(),
        "pinned_pages": pinned_pages,
    }
    pinned_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return pinned_pages
