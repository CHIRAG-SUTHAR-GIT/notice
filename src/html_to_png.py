"""Render local HTML files to PNG with Playwright's Chromium.

Run as a subprocess so Playwright's sync API never touches Streamlit's loop:

    python src/html_to_png.py jobs.json

jobs.json: {"width": 1240, "jobs": [{"html": "<path>", "png": "<path>"}, ...]}
"""

import json
from pathlib import Path
import sys

from playwright.sync_api import sync_playwright


def main() -> int:
    spec = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    width = int(spec.get("width", 1240))
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch()
        page = browser.new_page(viewport={"width": width, "height": 900}, device_scale_factor=2)
        for job in spec["jobs"]:
            page.goto(Path(job["html"]).resolve().as_uri())
            page.wait_for_load_state("networkidle")
            page.screenshot(path=job["png"], full_page=True)
        browser.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
