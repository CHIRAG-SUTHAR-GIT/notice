"""Sent-mail screenshots for the KYC notice record.

Renders each sent notice the way it looks in the mailbox - subject, sender,
recipients, date, the notice body and the attachment chips - and saves it as a
PNG that can be forwarded on its own.
"""

from __future__ import annotations

import base64
from datetime import datetime
from html import escape
import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile
from typing import Dict, List, Optional, Sequence, Tuple

RECORDS_DIR = Path(__file__).resolve().parent.parent / "KYC Mail Records"
_SHOT_WIDTH = 1240
_ATTACHMENT_STYLES = {
    ".pdf": ("#E2412C", "PDF"),
    ".xlsx": ("#1D8A4C", "XLS"),
}


def inline_cid_images(html: str, images: Dict[str, Tuple[str, bytes]]) -> str:
    """Replace cid: references with data URIs so the HTML renders on its own."""
    for cid, (subtype, data) in images.items():
        encoded = base64.b64encode(data).decode()
        html = html.replace(f"cid:{cid}", f"data:image/{subtype};base64,{encoded}")
    return html


def _attachment_chip(name: str, size_bytes: Optional[int]) -> str:
    color, label = _ATTACHMENT_STYLES.get(Path(name).suffix.lower(), ("#6B7280", "FILE"))
    size = f"{max(1, round((size_bytes or 0) / 1024))} KB" if size_bytes else ""
    return (
        '<div style="display:flex;align-items:center;gap:10px;border:1px solid #DADCE0;'
        'border-radius:6px;padding:8px 14px 8px 8px;background:#fff">'
        f'<div style="width:38px;height:38px;border-radius:4px;background:{color};color:#fff;'
        'font:700 11px Arial;display:flex;align-items:center;justify-content:center">'
        f"{label}</div>"
        f'<div style="font:13px Arial;color:#202124">{escape(name)}'
        f'<div style="font:12px Arial;color:#5F6368">{size}</div></div></div>'
    )


def build_sent_mail_html(
    subject: str,
    sender_name: str,
    sender_email: str,
    recipients: Sequence[str],
    sent_at: str,
    body_html: str,
    attachments: Sequence[Tuple[str, Optional[int]]],
) -> str:
    """Build the mailbox-style page for one sent notice."""
    try:
        stamp = datetime.strptime(sent_at, "%Y-%m-%d %H:%M:%S").strftime("%a, %d %b %Y %I:%M:%S %p +0530")
    except ValueError:
        stamp = sent_at
    to_line = ", ".join(f'"{address.split("@")[0]}" &lt;{escape(address)}&gt;' for address in recipients)
    chips = "".join(_attachment_chip(name, size) for name, size in attachments)
    sender_address = f" &lt;{escape(sender_email)}&gt;" if sender_email.strip() else ""
    count = len(attachments)

    return f"""<!doctype html><html><head><meta charset="utf-8"></head>
<body style="margin:0;background:#fff">
<div style="width:{_SHOT_WIDTH}px;padding:18px 24px;box-sizing:border-box;font-family:Arial,Helvetica,sans-serif">
  <div style="font:700 15px Arial;color:#202124;margin-bottom:14px">{escape(subject)}</div>
  <div style="display:flex;gap:12px;align-items:flex-start;border-bottom:1px solid #E8EAED;padding-bottom:12px">
    <div style="width:34px;height:34px;border-radius:4px;background:#A7D8DE;color:#0B3B45;
                font:700 13px Arial;display:flex;align-items:center;justify-content:center">Me</div>
    <div>
      <div style="font:13px Arial;color:#202124"><b>{escape(sender_name)}</b>{sender_address}</div>
      <div style="font:12px Arial;color:#5F6368;margin-top:3px">{escape(stamp)}</div>
      <div style="font:12px Arial;color:#5F6368;margin-top:3px">To {to_line}</div>
    </div>
  </div>
  <div style="padding:18px 6px 8px 6px">{body_html}</div>
  <div style="border-top:1px solid #E8EAED;padding-top:12px;margin-top:8px">
    <div style="font:13px Arial;color:#202124;margin-bottom:10px">{count} Attachment{"s" if count != 1 else ""}</div>
    <div style="display:flex;gap:14px;flex-wrap:wrap">{chips}</div>
  </div>
</div></body></html>"""


def render_html_to_png(html: str, output_path: Path, width: int = _SHOT_WIDTH) -> Path:
    """Render one HTML page to a PNG using the bundled Chromium, in a subprocess.

    A subprocess keeps Playwright's sync API away from Streamlit's event loop.
    """
    return render_html_files_to_png([(html, output_path)], width)[0]


def render_html_files_to_png(pages: List[Tuple[str, Path]], width: int = _SHOT_WIDTH) -> List[Path]:
    if not pages:
        return []
    with tempfile.TemporaryDirectory() as work_dir:
        jobs = []
        for index, (html, output_path) in enumerate(pages):
            source = Path(work_dir) / f"page_{index}.html"
            source.write_text(html, encoding="utf-8")
            output_path.parent.mkdir(parents=True, exist_ok=True)
            jobs.append({"html": str(source), "png": str(output_path)})
        job_file = Path(work_dir) / "jobs.json"
        job_file.write_text(json.dumps({"width": width, "jobs": jobs}), encoding="utf-8")
        result = subprocess.run(
            [sys.executable, str(Path(__file__).resolve().parent / "html_to_png.py"), str(job_file)],
            capture_output=True,
            text=True,
            timeout=120 + 20 * len(jobs),
        )
        if result.returncode != 0:
            raise RuntimeError(f"Screenshot failed: {result.stderr.strip() or result.stdout.strip()}")
    return [output for _, output in pages]


def screenshot_filename(bank_name: str) -> str:
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', "_", bank_name).strip(" ._") or "BANK"
    return f"{safe} - sent mail.png"


def records_folder(notice_date) -> Path:
    return RECORDS_DIR / str(notice_date)
