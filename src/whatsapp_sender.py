"""Send images on WhatsApp Web from this PC, without any extension.

The login is done once by scanning the QR code; Chromium keeps it in a profile
folder next to this project, so later runs reuse it.

    python src/whatsapp_sender.py login
    python src/whatsapp_sender.py send jobs.json
    python src/whatsapp_sender.py chats          # list chat names, for checking

jobs.json: {"chat": "<contact or group name>",
            "items": [{"image": "<png path>", "caption": "<text>"}, ...]}
Progress is printed as one JSON object per line.
"""

from __future__ import annotations

import base64
from datetime import datetime
import json
import mimetypes
import os
from pathlib import Path
import sys
import time
from typing import Dict, List

from playwright.sync_api import TimeoutError as PlaywrightTimeout, sync_playwright

_DATA_DIR = Path(__file__).resolve().parent.parent / "data"
# One folder per sender, so different officers can stay logged in side by side.
PROFILE_DIR = _DATA_DIR / os.environ.get("WHATSAPP_PROFILE", "whatsapp_profile")
DEBUG_DIR = _DATA_DIR / "whatsapp_debug"
WHATSAPP_URL = "https://web.whatsapp.com/"
# WhatsApp Web refuses headless browsers, so the window is always shown and the
# browser identifies itself as ordinary Chrome.
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36"
)
_CHAT_ROWS = '[role="grid"] [role="row"]'
_SEARCH_BOX = '[role="textbox"][data-tab="3"], input[aria-label*="Search" i], div[contenteditable="true"][data-tab="3"]'
_IMAGE_INPUT = 'input[type="file"][accept*="image"]'
_ATTACH_BUTTON = 'footer [aria-label="Attach"], footer [title="Attach"], [data-icon="ic-attach-file"]'
_CAPTION_BOX = '[role="textbox"][data-tab="10"], div[contenteditable="true"][data-tab="10"]'
_SEND_BUTTON = '[aria-label^="Send"], [data-icon="wds-ic-send-filled"], [data-icon="send"]'
_CLOSE_BUTTON = '[aria-label="Close"], [data-icon="x"], [data-icon="x-viewer"]'


def _emit(**payload) -> None:
    print(json.dumps(payload), flush=True)


def _debug_shot(page, label: str) -> str:
    """Save what the WhatsApp window looked like, to explain a failure."""
    try:
        DEBUG_DIR.mkdir(parents=True, exist_ok=True)
        path = DEBUG_DIR / f"{datetime.now():%Y%m%d_%H%M%S}_{label}.png"
        page.screenshot(path=str(path))
        return str(path)
    except Exception:
        return ""


def _open(playwright):
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    context = playwright.chromium.launch_persistent_context(
        str(PROFILE_DIR),
        headless=False,
        viewport={"width": 1280, "height": 860},
        user_agent=USER_AGENT,
        args=["--disable-blink-features=AutomationControlled"],
    )
    page = context.pages[0] if context.pages else context.new_page()
    page.goto(WHATSAPP_URL, wait_until="domcontentloaded")
    return context, page


def _wait_logged_in(page, timeout_ms: int) -> bool:
    try:
        page.wait_for_selector(_CHAT_ROWS, timeout=timeout_ms)
        return True
    except PlaywrightTimeout:
        return False


def login(timeout_seconds: int = 240) -> int:
    with sync_playwright() as playwright:
        context, page = _open(playwright)
        try:
            if _wait_logged_in(page, 15000):
                _emit(status="already_logged_in")
                return 0
            _emit(status="scan_qr", message="Scan the QR code in the WhatsApp window with your phone.")
            ok = _wait_logged_in(page, timeout_seconds * 1000)
            _emit(status="logged_in" if ok else "timeout")
            time.sleep(2)
            return 0 if ok else 1
        finally:
            context.close()


def _chat_titles(page) -> List[str]:
    return page.eval_on_selector_all(
        f'{_CHAT_ROWS} span[title]', "nodes => nodes.map(node => node.getAttribute('title'))"
    )


def _open_chat(page, chat: str) -> None:
    """Search for the chat by name and open the row that matches it."""
    search = page.locator(_SEARCH_BOX).first
    search.click()
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    search.type(chat, delay=40)
    page.wait_for_timeout(2500)

    rows = page.locator(_CHAT_ROWS)
    wanted = chat.strip().casefold()
    target = None
    for index in range(rows.count()):
        row = rows.nth(index)
        titles = row.locator("span[title]")
        if not titles.count():
            continue
        title = (titles.first.get_attribute("title") or "").strip()
        if title.casefold() == wanted:
            target = row
            break
        if target is None and wanted in title.casefold():
            target = row
    if target is None:
        found = ", ".join(_chat_titles(page)[:5])
        raise RuntimeError(
            f'WhatsApp chat "{chat}" was not found. Search showed: {found or "nothing"} '
            f"(debug shot: {_debug_shot(page, 'chat')})"
        )

    target.click()
    page.wait_for_selector("footer", timeout=20000)
    page.wait_for_timeout(1200)


def _dismiss_preview(page) -> None:
    for selector in _CLOSE_BUTTON.split(", "):
        control = page.locator(selector)
        if control.count():
            try:
                control.first.click()
                page.wait_for_timeout(800)
                return
            except Exception:
                pass
    page.keyboard.press("Escape")
    page.wait_for_timeout(800)


def _attach_with_file_input(page, image: Path) -> bool:
    file_input = page.locator(_IMAGE_INPUT).first
    if file_input.count() == 0:
        attach = page.locator(_ATTACH_BUTTON).first
        if attach.count() == 0:
            return False
        attach.click()
        page.wait_for_timeout(900)
        file_input = page.locator(_IMAGE_INPUT).first
        if file_input.count() == 0:
            return False
    file_input.set_input_files(str(image))
    return True


def _attach_with_paste(page, image: Path) -> bool:
    """Paste the image into the composer, like Ctrl+V does."""
    composer = page.locator(f"footer {_CAPTION_BOX}").first
    if composer.count() == 0:
        composer = page.locator(_CAPTION_BOX).first
    if composer.count() == 0:
        return False
    composer.click()
    payload = base64.b64encode(image.read_bytes()).decode()
    mime = mimetypes.guess_type(image.name)[0] or "image/png"
    page.evaluate(
        """({data, mime, name}) => {
            const bytes = Uint8Array.from(atob(data), c => c.charCodeAt(0));
            const file = new File([bytes], name, {type: mime});
            const transfer = new DataTransfer();
            transfer.items.add(file);
            const target = document.activeElement || document.body;
            target.dispatchEvent(new ClipboardEvent('paste', {
                clipboardData: transfer, bubbles: true, cancelable: true
            }));
        }""",
        {"data": payload, "mime": mime, "name": image.name},
    )
    return True


def _send_image(page, image: Path, caption: str) -> None:
    attached = False
    for strategy in (_attach_with_file_input, _attach_with_paste):
        try:
            if not strategy(page, image):
                continue
            page.wait_for_selector(_SEND_BUTTON, timeout=25000)
            attached = True
            break
        except PlaywrightTimeout:
            _dismiss_preview(page)
        except Exception:
            _dismiss_preview(page)
    if not attached:
        raise RuntimeError(f"Could not attach the image (debug shot: {_debug_shot(page, 'attach')})")

    if caption:
        caption_box = page.locator(_CAPTION_BOX).last
        caption_box.click()
        caption_box.type(caption, delay=12)
        page.wait_for_timeout(500)
    page.locator(_SEND_BUTTON).first.click()

    # The preview closes once the image is handed to the chat.
    for _ in range(40):
        page.wait_for_timeout(500)
        if page.locator(_SEND_BUTTON).count() == 0:
            return
    raise RuntimeError(
        f"The image preview did not close; WhatsApp may not have sent it "
        f"(debug shot: {_debug_shot(page, 'send')})"
    )


def send(job_file: str) -> int:
    spec: Dict = json.loads(Path(job_file).read_text(encoding="utf-8"))
    chat, items = spec["chat"], spec["items"]
    failures = 0
    with sync_playwright() as playwright:
        context, page = _open(playwright)
        try:
            if not _wait_logged_in(page, 60000):
                _emit(status="not_logged_in", message="Run the WhatsApp login first.",
                      debug=_debug_shot(page, "login"))
                return 1
            try:
                _open_chat(page, chat)
            except Exception as exc:
                _emit(status="chat_not_found", error=str(exc).splitlines()[0])
                return 1
            _emit(status="chat_open", chat=chat, total=len(items))
            for index, item in enumerate(items, start=1):
                caption = item.get("caption", "")
                try:
                    _send_image(page, Path(item["image"]), caption)
                    _emit(status="sent", index=index, caption=caption)
                    page.wait_for_timeout(1200)
                except Exception as exc:
                    failures += 1
                    _emit(status="failed", index=index, caption=caption, error=str(exc).splitlines()[0])
                    _dismiss_preview(page)
            _emit(status="done", failed=failures)
            time.sleep(2)
        finally:
            context.close()
    return 1 if failures else 0


def chats() -> int:
    """Print the chat names WhatsApp currently shows, to confirm exact spelling."""
    with sync_playwright() as playwright:
        context, page = _open(playwright)
        try:
            if not _wait_logged_in(page, 60000):
                _emit(status="not_logged_in")
                return 1
            _emit(status="chats", names=_chat_titles(page))
            return 0
        finally:
            context.close()


def main() -> int:
    if len(sys.argv) < 2 or sys.argv[1] not in {"login", "send", "chats"}:
        print(__doc__)
        return 2
    if sys.argv[1] == "login":
        return login()
    if sys.argv[1] == "chats":
        return chats()
    return send(sys.argv[2])


if __name__ == "__main__":
    sys.exit(main())
