from datetime import date
import email

import pandas as pd
import pytest

from src.bankwise_kyc_notices import build_bank_bundles, pending_rows
from src.kyc_notice_mailer import (
    build_message,
    build_subject,
    extract_emails,
    hash_page_password,
    load_access_record,
    save_access_record,
    verify_page_password,
    load_email_directory,
    mail_attachments,
    match_bank_emails,
    notice_docx_to_email,
    notice_subject_from_docx,
)


def _pending_sheet():
    return pd.DataFrame(
        [
            {
                "DATE": "01-09-2026",
                "ACK NO": "31108260205864",
                "IFSC CODE": "SBIN0060306",
                "BANK NAME": "State Bank of India",
                "AC NO": "32122008040",
                "ACCOUNT HOLDER'S NAME": "",
                "ACCOUNT HOLDER'S ADDRESS": "",
                "ACCOUNT HOLDER'S MOBILE NUMBER": "",
                "ACCOUNT HOLDER'S LOCATION": "",
            },
            {
                "DATE": "01-09-2026",
                "ACK NO": "31108260205865",
                "IFSC CODE": "UTIB0002362",
                "BANK NAME": "Axis Bank",
                "AC NO": "921010042651270",
                "ACCOUNT HOLDER'S NAME": "KNOWN HOLDER",
                "ACCOUNT HOLDER'S ADDRESS": "SURAT",
                "ACCOUNT HOLDER'S MOBILE NUMBER": "9999999999",
                "ACCOUNT HOLDER'S LOCATION": "",
            },
        ]
    )


@pytest.fixture(scope="module")
def sbi_bundle():
    bundles = build_bank_bundles(pending_rows(_pending_sheet()), date(2026, 9, 17))
    assert [bundle["bank_name"] for bundle in bundles] == ["State Bank of India"]
    return bundles[0]


def _docx(bundle):
    return next(data for name, data in bundle["files"].items() if name.endswith(".docx"))


def test_email_body_has_notice_text_but_no_signature_block(sbi_bundle):
    html, images = notice_docx_to_email(_docx(sbi_bundle))

    assert "The Nodal Officer," in html
    assert "State Bank of India" in html
    assert "No. CCoE/1930/DR/17/09/2026" in html
    assert "32122008040" in html
    assert "You are instructed to comply" in html
    assert "Keshvala" not in html
    assert "Superintendent of Police" not in html
    # Only the letterhead logo is kept; the stamp/signature image is dropped.
    assert len(images) == 1
    assert all(f"cid:{cid}" in html for cid in images)


def test_subject_uses_notice_subject_with_manual_period_and_bank(sbi_bundle):
    base = notice_subject_from_docx(_docx(sbi_bundle))

    assert build_subject(base, "September 1 to 15", "State Bank of India") == (
        "Provision of Know Your Customer (KYC) Details of Account Holder(s) "
        "(September 1 to 15- State Bank of India)"
    )


def test_message_attaches_only_pdf_and_excel(sbi_bundle):
    html, images = notice_docx_to_email(_docx(sbi_bundle))
    attachments = mail_attachments(sbi_bundle)
    message = build_message(
        "sender@gujarat.gov.in", ["nodal@sbi.co.in"], "Subject", html, images, attachments, cc=["cc@gujarat.gov.in"]
    )

    parsed = email.message_from_bytes(message.as_bytes())
    filenames = sorted(part.get_filename() for part in parsed.walk() if part.get_filename())
    assert filenames == ["State Bank of India KYC Details.xlsx", "State Bank of India KYC Notice.pdf"]
    assert parsed["To"] == "nodal@sbi.co.in"
    assert parsed["Cc"] == "cc@gujarat.gov.in"


def test_build_message_refuses_word_attachments():
    with pytest.raises(ValueError, match="only PDF and Excel"):
        build_message("a@b.in", ["c@d.in"], "S", "<p>x</p>", {}, {"notice.docx": b"x"})


def test_extract_emails_splits_and_dedupes():
    assert extract_emails("a@sbi.co.in; B@sbi.co.in,\na@SBI.co.in  c@x.bank.in") == [
        "a@sbi.co.in",
        "B@sbi.co.in",
        "c@x.bank.in",
    ]


def test_match_bank_emails_exact_fuzzy_and_rejections():
    directory = load_email_directory(
        pd.DataFrame(
            {
                "CATEGORY": [
                    "STATE BANK OF INDIA",
                    "STATE BANK OF INDIA UPI",
                    "AU SMALL FINANCE BANK UPI",
                    "AU SMALL FINANCE BANK",
                    "THE VARACHHAA CO-OPERATIV BANK LTD",
                    "UTTAR PRADESH COOPERATIVE BANK LTD",
                    "UTTAR PRADESH POWER CORPORATION LIMITED",
                    "EQUITAS BANK UPI",
                ],
                "MAIL IDS": ["a@sbi.co.in", "b@sbi.co.in", "upi@au.in", "c@au.in", "d@varachha.in", "e@upcb.in", "f@uppcl.org", "g@equitas.in"],
            }
        ),
        "CATEGORY",
        "MAIL IDS",
    )

    sbi = match_bank_emails("State Bank of India", directory)
    assert sbi["match"] == "Exact"
    assert sbi == {"matched_name": "STATE BANK OF INDIA", "emails": ["a@sbi.co.in"], "match": "Exact"}
    upi_only = match_bank_emails("Equitas Bank", directory)
    assert upi_only == {"matched_name": "EQUITAS BANK UPI", "emails": ["g@equitas.in"], "match": "Exact"}
    au = match_bank_emails("AU Bank", directory)
    assert au == {"matched_name": "AU SMALL FINANCE BANK", "emails": ["c@au.in"], "match": "Exact"}
    assert match_bank_emails("THE VARACHHA CO.OP.BANK LTD.", directory)["emails"] == ["d@varachha.in"]
    assert match_bank_emails("Uttar Pradesh Gramin Bank", directory)["match"] == "Not found"


def test_page_password_is_stored_hashed_and_verified(tmp_path):
    access_file = tmp_path / "access.json"
    save_access_record("Strong#Pass1", access_file)

    record = load_access_record(access_file)
    assert "Strong#Pass1" not in access_file.read_text()
    assert verify_page_password("Strong#Pass1", record)
    assert not verify_page_password("strong#pass1", record)
    assert hash_page_password("Strong#Pass1")["salt"] != record["salt"]
    assert load_access_record(tmp_path / "missing.json") is None


def test_sent_mail_screenshot_html_and_png(tmp_path, sbi_bundle):
    from src.mail_screenshot import build_sent_mail_html, render_html_files_to_png

    html = build_sent_mail_html(
        "Subject (September 1 to 15- State Bank of India)",
        "HELPLINE UNIT-16",
        "helpline@example.gov.in",
        ["nodal@sbi.co.in"],
        "2026-09-17 19:37:18",
        "<p>notice body</p>",
        [("State Bank of India KYC Notice.pdf", 79000), ("State Bank of India KYC Details.xlsx", 6100)],
    )

    assert "September 1 to 15- State Bank of India" in html
    assert "nodal@sbi.co.in" in html and "helpline@example.gov.in" in html
    assert "Thu, 17 Sep 2026 07:37:18 PM +0530" in html
    assert "KYC Notice.pdf" in html and "KYC Details.xlsx" in html and "2 Attachments" in html
    assert "Keshvala" not in html

    png = tmp_path / "shot.png"
    render_html_files_to_png([(html, png)])
    assert png.exists() and png.stat().st_size > 5000
