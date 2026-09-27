"""CSV fixer page and processor.

Ports the standalone CSV Fixer folder logic into the Streamlit app.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

import streamlit as st


EXPECTED_COLUMNS = 16
NUM_TRAILING_COLUMNS = 14


@dataclass(frozen=True)
class CsvFixStats:
    """Summary of the CSV repair pass."""

    total_lines: int
    data_lines: int
    fixed_lines: int
    unchanged_lines: int
    blank_lines: int


def _decode_uploaded_csv(raw_bytes: bytes) -> str:
    """Decode uploaded CSV bytes with sensible fallbacks."""
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return raw_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw_bytes.decode("utf-8", errors="replace")


def fix_csv_text(csv_text: str) -> tuple[str, CsvFixStats]:
    """Fix raw commas in the second column using the standalone tool logic."""
    lines = re.split(r"\r?\n", csv_text)
    fixed_lines: list[str] = []
    fixed_count = 0
    blank_count = 0
    unchanged_count = 0

    for line in lines:
        if not line.strip():
            fixed_lines.append(line)
            blank_count += 1
            continue

        parts = line.split(",")

        if len(parts) > EXPECTED_COLUMNS:
            s_no = parts[0]
            trailing_start_index = len(parts) - NUM_TRAILING_COLUMNS
            trailing_columns = parts[trailing_start_index:]
            name_parts = parts[1:trailing_start_index]
            fixed_name = f'"{",".join(name_parts)}"'
            fixed_lines.append(",".join([s_no, fixed_name, *trailing_columns]))
            fixed_count += 1
        else:
            fixed_lines.append(line)
            unchanged_count += 1

    stats = CsvFixStats(
        total_lines=len(lines),
        data_lines=len(lines) - blank_count,
        fixed_lines=fixed_count,
        unchanged_lines=unchanged_count,
        blank_lines=blank_count,
    )
    return "\n".join(fixed_lines), stats


def _fixed_filename(original_name: str) -> str:
    """Create the same download naming pattern as the standalone tool."""
    return f"fixed_{original_name or 'file.csv'}"


def render_csv_fixer_page():
    """Render the CSV Fixer Streamlit page."""
    st.title("CSV Fixer")
    st.markdown(
        "Upload a corrupted CSV file and fix misplaced commas in the bank/name column "
        "so all rows align back to the expected 16-column structure."
    )

    if "csv_fixer_upload_nonce" not in st.session_state:
        st.session_state.csv_fixer_upload_nonce = 0

    uploaded_file = st.file_uploader(
        "Choose a CSV file",
        type=["csv"],
        key=f"csv_fixer_uploader_{st.session_state.csv_fixer_upload_nonce}",
        help="Use this for CSVs where the second column contains raw commas without quotes.",
    )

    with st.expander("How it fixes the CSV", expanded=False):
        st.markdown(
            """
            - Expects 16 columns per row.
            - Treats the final 14 columns as fixed trailing data.
            - If a row has extra commas, everything between column 1 and those final 14 columns is joined back into column 2.
            - The repaired column 2 value is wrapped in quotes, matching the standalone CSV Fixer logic.
            """
        )

    if uploaded_file is None:
        st.info("Upload a CSV file to start fixing it.")
        return

    if not uploaded_file.name.lower().endswith(".csv"):
        st.error("Please upload a valid CSV file.")
        return

    try:
        original_text = _decode_uploaded_csv(uploaded_file.getvalue())
        fixed_text, stats = fix_csv_text(original_text)
    except Exception as exc:  # pragma: no cover - defensive UI guard
        st.error("Failed to process the CSV file. It might be too corrupted or in an unknown format.")
        st.exception(exc)
        return

    st.success("Your CSV file has been successfully fixed and aligned.")

    metric_cols = st.columns(4)
    metric_cols[0].metric("Lines scanned", f"{stats.data_lines:,}")
    metric_cols[1].metric("Rows fixed", f"{stats.fixed_lines:,}")
    metric_cols[2].metric("Rows unchanged", f"{stats.unchanged_lines:,}")
    metric_cols[3].metric("Blank lines", f"{stats.blank_lines:,}")

    download_col, reset_col = st.columns([2, 1])
    with download_col:
        st.download_button(
            "Download Fixed CSV",
            data=fixed_text.encode("utf-8"),
            file_name=_fixed_filename(uploaded_file.name),
            mime="text/csv",
            type="primary",
            use_container_width=True,
            key="csv_fixer_download",
        )
    with reset_col:
        if st.button("Fix Another File", use_container_width=True, key="csv_fixer_reset"):
            st.session_state.csv_fixer_upload_nonce += 1
            st.rerun()

    with st.expander("Preview fixed CSV lines", expanded=False):
        preview_lines = fixed_text.splitlines()[:20]
        if preview_lines:
            st.code("\n".join(preview_lines), language="csv")
        else:
            st.info("The fixed file is empty.")
