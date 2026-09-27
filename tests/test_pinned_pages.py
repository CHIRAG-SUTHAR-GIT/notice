import json

from src.pinned_pages import MAX_PINNED_PAGES, load_pinned_pages, normalize_pinned_pages, save_pinned_pages


def test_normalize_pinned_pages_keeps_valid_unique_pages_in_order():
    valid_pages = ["upload", "csv_fixer", "mo_finder"]

    result = normalize_pinned_pages(
        ["mo_finder", "missing", "csv_fixer", "mo_finder", "upload"],
        valid_pages,
    )

    assert result == ["mo_finder", "csv_fixer", "upload"]


def test_save_and_load_pinned_pages_persists_local_file(tmp_path):
    valid_pages = [f"page_{index}" for index in range(12)]
    path = tmp_path / "pinned_pages.json"

    saved = save_pinned_pages(
        [*valid_pages, "unknown", valid_pages[0]],
        valid_pages,
        path=path,
    )

    assert saved == valid_pages[:MAX_PINNED_PAGES]
    assert load_pinned_pages(valid_pages, path=path) == valid_pages[:MAX_PINNED_PAGES]

    payload = json.loads(path.read_text(encoding="utf-8"))
    assert payload["version"] == 1
    assert payload["pinned_pages"] == valid_pages[:MAX_PINNED_PAGES]


def test_load_pinned_pages_ignores_broken_json(tmp_path):
    path = tmp_path / "pinned_pages.json"
    path.write_text("{broken", encoding="utf-8")

    assert load_pinned_pages(["upload"], path=path) == []
