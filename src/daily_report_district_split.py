"""Daily Report District Split page."""

from __future__ import annotations

import re
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from src.daily_report_district_engine import (
    ALLOWED_EXTENSIONS,
    IFSC_COLUMN_CANDIDATES,
    MAX_FILE_SIZE,
    PINCODE_COLUMN_CANDIDATES,
    create_master_zip,
    find_column,
    process_input_file,
    read_input_file,
)


@dataclass(frozen=True)
class DailyReportDistrictSplitResult:
    """In-memory output from the daily report district splitter."""

    workbook_zip: bytes
    filename: str
    summary: dict[str, Any]
    unmatched: list[dict[str, Any]]


def _safe_upload_name(filename: str) -> str:
    path = Path(filename or "daily_report.xlsx")
    suffix = path.suffix.lower()
    stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", path.stem).strip("._") or "daily_report"
    return f"{stem}{suffix}"


def preview_daily_report_file(raw_bytes: bytes, filename: str) -> dict[str, Any]:
    """Read uploaded bytes into a temporary source file and return preview metadata."""
    suffix = Path(filename or "").suffix.lower().lstrip(".")
    if suffix not in ALLOWED_EXTENSIONS:
        raise ValueError("Upload an XLSX, XLS, or CSV file.")
    if len(raw_bytes) > MAX_FILE_SIZE:
        raise ValueError("File too large. Maximum size is 50MB.")

    with tempfile.TemporaryDirectory(prefix="datalens_daily_district_preview_") as temp_dir:
        source_path = Path(temp_dir) / _safe_upload_name(filename)
        source_path.write_bytes(raw_bytes)
        df = read_input_file(source_path)

    columns = list(df.columns)
    return {
        "columns": columns,
        "row_count": int(len(df)),
        "suggested_ifsc_column": find_column(columns, IFSC_COLUMN_CANDIDATES),
        "suggested_pincode_column": find_column(columns, PINCODE_COLUMN_CANDIDATES),
        "sample_rows": df.head(10).fillna("").astype(str),
    }


def process_daily_report_district_split(
    raw_bytes: bytes,
    filename: str,
    ifsc_column: str | None = "__auto__",
    pincode_column: str | None = "__auto__",
) -> DailyReportDistrictSplitResult:
    """Run the standalone daily report district split logic without persistent files."""
    suffix = Path(filename or "").suffix.lower().lstrip(".")
    if suffix not in ALLOWED_EXTENSIONS:
        raise ValueError("Upload an XLSX, XLS, or CSV file.")
    if len(raw_bytes) > MAX_FILE_SIZE:
        raise ValueError("File too large. Maximum size is 50MB.")

    with tempfile.TemporaryDirectory(prefix="datalens_daily_district_split_") as temp_dir:
        source_path = Path(temp_dir) / _safe_upload_name(filename)
        source_path.write_bytes(raw_bytes)
        district_data, unmatched, summary = process_input_file(
            source_path,
            ifsc_column=ifsc_column,
            pincode_column=pincode_column,
        )
        zip_bytes = create_master_zip(district_data, unmatched, summary)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return DailyReportDistrictSplitResult(
        workbook_zip=zip_bytes,
        filename=f"Daily_Report_District_Split_{timestamp}.zip",
        summary=summary,
        unmatched=unmatched,
    )


def _column_selection_value(label: str, columns: list[str], suggested: str | None, key: str) -> str:
    options = ["Auto detect", "Do not use", *columns]
    default_index = 0
    if suggested in columns:
        default_index = options.index(suggested)
    selected = st.selectbox(label, options, index=default_index, key=key)
    if selected == "Auto detect":
        return "__auto__"
    if selected == "Do not use":
        return "__none__"
    return selected


def render_daily_report_district_split_page():
    """Render the Daily Report District Split page."""
    st.title("Dailly Reports Split Suspect District Gujarat")
    st.markdown(
        "Upload a daily report file and split it into Gujarat district-wise Excel files using "
        "IFSC lookup first, then pincode fallback."
    )

    uploaded_file = st.file_uploader(
        "Choose daily report file",
        type=sorted(ALLOWED_EXTENSIONS),
        key="daily_report_district_split_uploader",
        help="Supported formats: XLSX, XLS, CSV. The file should contain IFSC and/or pincode columns.",
    )

    if uploaded_file is None:
        st.info("Upload a daily report to preview columns and generate the district-wise ZIP.")
        return

    raw_bytes = uploaded_file.getvalue()
    try:
        preview = preview_daily_report_file(raw_bytes, uploaded_file.name)
    except Exception as exc:
        st.error(str(exc))
        return

    st.success(f"Loaded {preview['row_count']:,} rows from {uploaded_file.name}.")
    columns = preview["columns"]
    col_a, col_b = st.columns(2)
    with col_a:
        ifsc_column = _column_selection_value(
            "IFSC column",
            columns,
            preview["suggested_ifsc_column"],
            "daily_report_district_ifsc_column",
        )
    with col_b:
        pincode_column = _column_selection_value(
            "Pincode column",
            columns,
            preview["suggested_pincode_column"],
            "daily_report_district_pincode_column",
        )

    with st.expander("Preview uploaded rows", expanded=False):
        st.dataframe(preview["sample_rows"], use_container_width=True, hide_index=True)

    with st.expander("What this page creates", expanded=False):
        st.markdown(
            """
            - A ZIP containing one Excel file per matched Gujarat district.
            - An **Other.xlsx** file for rows where no Gujarat district can be determined.
            - Daily-report columns are preserved in a clean output order with styled headers.
            """
        )

    if st.button("Generate Daily District ZIP", type="primary", use_container_width=True):
        try:
            with st.spinner("Resolving districts and building ZIP..."):
                result = process_daily_report_district_split(
                    raw_bytes,
                    uploaded_file.name,
                    ifsc_column=ifsc_column,
                    pincode_column=pincode_column,
                )
        except Exception as exc:
            st.error(str(exc))
            return

        summary = result.summary
        st.success("Daily report district split completed.")
        metrics = st.columns(6)
        metrics[0].metric("Input rows", f"{summary.get('input_rows', 0):,}")
        metrics[1].metric("Matched rows", f"{summary.get('matched_rows', 0):,}")
        metrics[2].metric("Other rows", f"{summary.get('other_rows', 0):,}")
        metrics[3].metric("Districts", f"{summary.get('total_districts', 0):,}")
        metrics[4].metric("IFSC matches", f"{summary.get('matched_by_ifsc', 0):,}")
        metrics[5].metric("PIN matches", f"{summary.get('matched_by_pincode', 0):,}")

        district_rows = [
            {"District": district, "Rows": data.get("record_count", 0)}
            for district, data in summary.get("districts", {}).items()
        ]
        if district_rows:
            st.subheader("District Output")
            st.dataframe(pd.DataFrame(district_rows), use_container_width=True, hide_index=True)

        if result.unmatched:
            with st.expander(f"Unmatched rows ({len(result.unmatched):,})", expanded=False):
                st.dataframe(pd.DataFrame(result.unmatched[:50]), use_container_width=True, hide_index=True)

        st.download_button(
            "Download Daily District ZIP",
            data=result.workbook_zip,
            file_name=result.filename,
            mime="application/zip",
            type="primary",
            use_container_width=True,
            key="daily_report_district_download",
        )
