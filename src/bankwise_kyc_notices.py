"""Generate bank-wise KYC notices (Word, PDF and Excel) from a pending KYC sheet.

Each bank gets its own folder containing a notice addressed to that bank's
nodal officer and an Excel file listing the accounts whose KYC is pending.

Usage:
    python src/bankwise_kyc_notices.py "SEP PANDING KYC.xlsx"
    python src/bankwise_kyc_notices.py "SEP PANDING KYC.xlsx" --date 17-09-2026 --output "KYC Notices"
"""

import argparse
import re
from datetime import date, datetime
from pathlib import Path
import sys
from typing import Dict, List

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.gujarat_account_formatter import (  # noqa: E402
    _account_key,
    _clean_identifier,
    _clean_text,
    _current_notice_date,
    _safe_filename,
    build_kyc_notice_docx,
    build_kyc_notice_pdf,
    build_payer_remitter_docx,
    build_payer_remitter_pdf,
    dataframe_to_styled_excel_bytes,
    prepare_kyc_notice_accounts,
)

KYC_DETAIL_COLUMNS = [
    "ACCOUNT HOLDER'S NAME",
    "ACCOUNT HOLDER'S MOBILE NUMBER",
    "ACCOUNT HOLDER'S ADDRESS",
    "ACCOUNT HOLDER'S LOCATION",
]
EXCEL_COLUMNS = ["SR NO", "ACK NO", "IFSC CODE", "BANK NAME", "AC NO", *KYC_DETAIL_COLUMNS]


def pending_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Keep rows missing a name, address or mobile number."""
    df = df.copy()
    for column in ["ACK NO", "IFSC CODE", "BANK NAME", "AC NO", *KYC_DETAIL_COLUMNS]:
        if column not in df.columns:
            df[column] = ""
    required = KYC_DETAIL_COLUMNS[:3]
    pending_mask = df[required].apply(lambda col: col.map(_clean_text) == "").any(axis=1)
    return df[pending_mask].reset_index(drop=True)


def load_pending_rows(path) -> pd.DataFrame:
    return pending_rows(pd.read_excel(path, dtype=str, keep_default_na=False))


def bank_excel_rows(bank_rows: pd.DataFrame, notice_df: pd.DataFrame) -> pd.DataFrame:
    """Build the Excel rows in the same order and numbering as the notice table."""
    by_account = {}
    for record in bank_rows.to_dict("records"):
        by_account.setdefault(_account_key(record["AC NO"]), record)

    rows = []
    for notice in notice_df.itertuples(index=False):
        source = by_account[_account_key(notice[2])]
        rows.append(
            {
                "SR NO": notice[0],
                "ACK NO": notice[1],
                "IFSC CODE": _clean_text(source["IFSC CODE"]),
                "BANK NAME": notice[3],
                "AC NO": notice[2],
                **{column: _clean_identifier(source[column]) for column in KYC_DETAIL_COLUMNS},
            }
        )
    return pd.DataFrame(rows, columns=EXCEL_COLUMNS)


def bank_folder_name(bank_name: str) -> str:
    # Parenthetical notes like "(including Andhra Bank ...)" keep paths under Windows' limit.
    return _safe_filename(re.sub(r"\s*\(.*?\)", "", bank_name)) or _safe_filename(bank_name)


def build_bank_bundles(pending: pd.DataFrame, notice_date: date) -> List[Dict]:
    """Build the notice (Word + PDF) and KYC Excel for every bank in memory."""
    pending = pending.assign(
        _bank=pending["BANK NAME"].map(_clean_text).replace("", "UNKNOWN BANK")
    )
    bundles = []
    for bank_name, bank_rows in pending.groupby("_bank", sort=True):
        notice_df, stats = prepare_kyc_notice_accounts(bank_rows)
        name = bank_folder_name(bank_name)
        bundle = {
            "bank_name": bank_name,
            "folder_name": name,
            "accounts": len(notice_df),
            "skipped_rows": stats["invalid_account_rows"] + stats["duplicate_account_rows"],
        }
        if not notice_df.empty:
            bundle["files"] = {
                f"{name} KYC Notice.docx": build_kyc_notice_docx(notice_df, notice_date, bank_name=bank_name),
                f"{name} KYC Notice.pdf": build_kyc_notice_pdf(notice_df, notice_date, bank_name=bank_name),
                f"{name} KYC Details.xlsx": dataframe_to_styled_excel_bytes(
                    bank_excel_rows(bank_rows, notice_df), sheet_name="KYC Details"
                ),
            }
        bundles.append(bundle)
    return bundles


def build_payer_remitter_bundles(pending: pd.DataFrame, notice_date: date) -> List[Dict]:
    """Build the Payer/Remitter notice (Word + PDF) and Excel for every bank in memory."""
    pending = pending.assign(
        _bank=pending["BANK NAME"].map(_clean_text).replace("", "UNKNOWN BANK")
    )
    bundles = []
    for bank_name, bank_rows in pending.groupby("_bank", sort=True):
        notice_df, stats = prepare_kyc_notice_accounts(bank_rows)
        name = bank_folder_name(bank_name)
        bundle = {
            "bank_name": bank_name,
            "folder_name": name,
            "accounts": len(notice_df),
            "skipped_rows": stats["invalid_account_rows"] + stats["duplicate_account_rows"],
        }
        if not notice_df.empty:
            bundle["files"] = {
                f"{name} Payer Remitter Notice.docx": build_payer_remitter_docx(notice_df, notice_date, bank_name=bank_name),
                f"{name} Payer Remitter Notice.pdf": build_payer_remitter_pdf(notice_df, notice_date, bank_name=bank_name),
                f"{name} KYC Details.xlsx": dataframe_to_styled_excel_bytes(
                    bank_excel_rows(bank_rows, notice_df), sheet_name="KYC Details"
                ),
            }
        bundles.append(bundle)
    return bundles


def save_bundles(bundles: List[Dict], output_dir: Path) -> None:
    for bundle in bundles:
        if not bundle.get("files"):
            continue
        bank_dir = output_dir / bundle["folder_name"]
        bank_dir.mkdir(parents=True, exist_ok=True)
        for filename, data in bundle["files"].items():
            (bank_dir / filename).write_bytes(data)


def generate(source: Path, output_dir: Path, notice_date: date) -> None:
    bundles = build_bank_bundles(load_pending_rows(source), notice_date)
    save_bundles(bundles, output_dir)
    for bundle in bundles:
        if not bundle.get("files"):
            print(f"SKIPPED  {bundle['bank_name']}: no valid account numbers")
            continue
        dropped = bundle["skipped_rows"]
        note = f" ({dropped} duplicate/invalid row(s) skipped)" if dropped else ""
        print(f"{bundle['accounts']:>4} accounts  {bundle['bank_name']}{note}")
    total = sum(bundle["accounts"] for bundle in bundles)
    print(f"\n{len(bundles)} banks, {total} accounts -> {output_dir}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("source", type=Path, help="Pending KYC Excel file")
    parser.add_argument("--output", type=Path, help="Output folder (default: '<source name> Bank Notices')")
    parser.add_argument("--date", help="Notice date as DD-MM-YYYY (default: today)")
    args = parser.parse_args()

    notice_date = datetime.strptime(args.date, "%d-%m-%Y").date() if args.date else _current_notice_date()
    output_dir = args.output or args.source.with_name(f"{args.source.stem} Bank Notices")
    generate(args.source, output_dir, notice_date)


if __name__ == "__main__":
    main()
