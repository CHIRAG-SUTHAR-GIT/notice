"""MO Finder page.

Integrates the standalone MO FINDER report builder into the Streamlit app while
keeping uploads and generated workbooks in temporary storage only.
"""

from __future__ import annotations

import re
import tempfile
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
import streamlit as st

from src.mo_report_builder import (
    DEFAULT_ACK_PREFIX,
    SUPPORTED_SOURCE_EXTENSIONS,
    build_report,
)


@dataclass(frozen=True)
class MOFinderResult:
    """In-memory MO report payload and UI summary."""

    workbook: bytes
    filename: str
    source_file: str
    total_rows: int
    rows_counted: int
    rows_ignored: int
    unique_mo_categories: int
    top_10: list[dict[str, int | str]]


def _safe_upload_name(filename: str) -> str:
    """Create a filesystem-safe temporary upload name."""
    path = Path(filename or "upload.csv")
    suffix = path.suffix.lower()
    stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", path.stem).strip("._") or "upload"
    return f"{stem}{suffix}"


def process_mo_report_upload(raw_bytes: bytes, filename: str, acknowledgement_prefix: str) -> MOFinderResult:
    """Build the MO Unique Report using the copied standalone report engine."""
    suffix = Path(filename or "").suffix.lower()
    if suffix not in SUPPORTED_SOURCE_EXTENSIONS:
        raise ValueError("Upload a CSV, XLSX, or XLSM file.")

    with tempfile.TemporaryDirectory(prefix="datalens_mo_finder_") as temp_dir:
        temp_path = Path(temp_dir)
        source_path = temp_path / _safe_upload_name(filename)
        source_path.write_bytes(raw_bytes)

        result = build_report(
            source_file=source_path,
            output_dir=temp_path,
            acknowledgement_prefix=acknowledgement_prefix,
        )

        output_path = Path(result["output_file"])
        workbook = output_path.read_bytes()
        download_name = output_path.name

    return MOFinderResult(
        workbook=workbook,
        filename=download_name,
        source_file=str(result["source_file"]),
        total_rows=int(result["total_rows"]),
        rows_counted=int(result["rows_counted"]),
        rows_ignored=int(result["rows_ignored"]),
        unique_mo_categories=int(result["unique_mo_categories"]),
        top_10=list(result["top_10"]),
    )


def render_mo_finder_page():
    """Render the MO Finder page."""
    st.title("MO Finder")
    st.markdown(
        "Upload an Additional Information report to generate the same MO Unique Report workbook "
        "with grouped MO categories, detail mapping, rules, and top MO counts."
    )

    uploaded_file = st.file_uploader(
        "Choose a CSV, XLSX, or XLSM file",
        type=["csv", "xlsx", "xlsm"],
        key="mo_finder_uploader",
        help="Source file must include Acknowledgement No. and Crime Aditional Information columns.",
    )

    ack_prefix = st.text_input(
        "Acknowledgement prefix",
        value=DEFAULT_ACK_PREFIX,
        key="mo_finder_ack_prefix",
        help="Only acknowledgement numbers starting with this prefix are counted. Leave blank to include all rows.",
    ).strip()

    with st.expander("What this tool creates", expanded=False):
        st.markdown(
            """
            - **MO Summary:** unique MO categories, counts, percentages, variants, amount totals, and top chart.
            - **Detail Mapping:** acknowledgement number, original MO text, cleaned MO text, final grouped MO, category fields, status, and amount.
            - **Rules:** audit notes describing filtering, cleaning, and grouping rules.
            """
        )

    if uploaded_file is None:
        st.info("Upload a report file to generate the MO Unique Report.")
        return

    try:
        result = process_mo_report_upload(
            raw_bytes=uploaded_file.getvalue(),
            filename=uploaded_file.name,
            acknowledgement_prefix=ack_prefix,
        )
    except Exception as exc:
        st.error(str(exc))
        return

    st.success("MO Unique Report generated successfully.")

    metric_cols = st.columns(4)
    metric_cols[0].metric("Total rows", f"{result.total_rows:,}")
    metric_cols[1].metric("Rows counted", f"{result.rows_counted:,}")
    metric_cols[2].metric("Rows ignored", f"{result.rows_ignored:,}")
    metric_cols[3].metric("Unique MO", f"{result.unique_mo_categories:,}")

    st.download_button(
        "Download MO Report",
        data=result.workbook,
        file_name=result.filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
        key="mo_finder_download",
    )

    if result.top_10:
        st.subheader("Top MO Categories")
        st.dataframe(pd.DataFrame(result.top_10), use_container_width=True, hide_index=True)
