import io

import pandas as pd

from src.drop_call_finder import generate_filename_from_dates, process_drop_call_workbook


def _sample_excel_file() -> io.BytesIO:
    df = pd.DataFrame(
        [
            {"phone_number_dialed": "111", "status": "DROP", "call_date": "20-02-2026 20:01:00"},
            {"phone_number_dialed": "111", "status": "DROP", "call_date": "20-02-2026 20:05:00"},
            {"phone_number_dialed": "222", "status": "DROP", "call_date": "21-02-2026 19:59:00"},
            {"phone_number_dialed": "222", "status": "INCALL", "call_date": "21-02-2026 20:00:00"},
            {"phone_number_dialed": "333", "status": "FINFR", "call_date": "21-02-2026 20:01:00"},
            {"phone_number_dialed": "444", "status": "DROP", "call_date": "21-02-2026 20:02:00"},
        ]
    )
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False)
    output.seek(0)
    return output


def test_generate_filename_from_dates_uses_call_date_range():
    df = pd.DataFrame(
        {
            "call_date": [
                "21-02-2026 20:00:00",
                "20-02-2026 20:01:00",
            ]
        }
    )

    assert generate_filename_from_dates(df) == "20-02-2026 to 21-02-2026 call.xlsx"


def test_process_drop_call_workbook_creates_expected_sheets():
    result = process_drop_call_workbook(_sample_excel_file())

    assert result.filename == "20-02-2026 to 21-02-2026 call.xlsx"
    assert result.original_rows == 6
    assert result.unique_phone_numbers == 4
    assert result.drop_only_numbers == 2

    workbook = io.BytesIO(result.workbook)
    sheet2 = pd.read_excel(workbook, sheet_name="Sheet2")
    workbook.seek(0)
    sheet3 = pd.read_excel(workbook, sheet_name="Sheet3")
    workbook.seek(0)
    sheet4 = pd.read_excel(workbook, sheet_name="Sheet4")

    assert "Grand Total" in sheet2.columns
    assert set(sheet3["phone_number_dialed"].astype(str)) == {"111", "444"}
    assert sheet4["sr"].tolist() == [1, 2]
    assert set(sheet4["phone_number_dialed"].astype(str)) == {"111", "444"}
