import io
import zipfile

from openpyxl import load_workbook

import src.daily_report_district_engine as engine
from src.daily_report_district_split import (
    preview_daily_report_file,
    process_daily_report_district_split,
)


def _sample_daily_report_csv() -> bytes:
    return (
        "S No.,Acknowledgement No.,District,Account No.,IFSC,Address,Pincode,"
        "Transaction Amount,Disputed Amount,Bank/FIs,Layers,Victim District,Reported Amount (Victim)\n"
        "10,ACK001,,11111111111,GOODIFSC,Ahmedabad Branch,,1000,500,Bank A,1,Ahmedabad,1000\n"
        "20,ACK004,,44444444444,GOODIFSC,Ahmedabad Branch,,4000,2000,Bank A,1,Ahmedabad,4000\n"
        "30,ACK002,,22222222222,,Surat Branch,395001,2000,1200,Bank B,2,Surat,2000\n"
        "40,ACK003,,33333333333,,Unknown Branch,,3000,1500,Bank C,3,Rajkot,3000\n"
    ).encode("utf-8")


def test_preview_daily_report_file_detects_ifsc_and_pincode_columns():
    preview = preview_daily_report_file(_sample_daily_report_csv(), "daily.csv")

    assert preview["row_count"] == 4
    assert preview["suggested_ifsc_column"] == "IFSC"
    assert preview["suggested_pincode_column"] == "Pincode"


def test_process_daily_report_district_split_creates_district_zip(monkeypatch):
    def fake_resolve_row(ifsc_code, pincode):
        if ifsc_code == "GOODIFSC":
            return {"ok": True, "district": "Ahmedabad", "source": "IFSC"}
        if pincode == "395001":
            return {"ok": True, "district": "Surat", "source": "Pincode"}
        return {
            "ok": False,
            "district": None,
            "source": None,
            "reason": "Could not determine district",
        }

    monkeypatch.setattr(engine, "resolve_row", fake_resolve_row)

    result = process_daily_report_district_split(_sample_daily_report_csv(), "daily.csv")

    assert result.filename.startswith("Daily_Report_District_Split_")
    assert result.summary["input_rows"] == 4
    assert result.summary["matched_rows"] == 3
    assert result.summary["other_rows"] == 1
    assert result.summary["total_districts"] == 2
    assert result.summary["matched_by_ifsc"] == 2
    assert result.summary["matched_by_pincode"] == 1
    assert len(result.unmatched) == 1

    with zipfile.ZipFile(io.BytesIO(result.workbook_zip)) as archive:
        assert sorted(archive.namelist()) == ["AHMEDABAD.xlsx", "OTHER.xlsx", "SURAT.xlsx"]
        ahmedabad_wb = load_workbook(io.BytesIO(archive.read("AHMEDABAD.xlsx")), data_only=True)
        surat_wb = load_workbook(io.BytesIO(archive.read("SURAT.xlsx")), data_only=True)
        other_wb = load_workbook(io.BytesIO(archive.read("OTHER.xlsx")), data_only=True)

    try:
        sheet = ahmedabad_wb["Data"]
        headers = [cell.value for cell in sheet[1]]
        assert headers == engine.OUTPUT_COLUMNS
        assert [sheet["A2"].value, sheet["A3"].value] == [1, 2]
        assert sheet["B2"].value == "ACK001"
        assert sheet["C2"].value in (None, "")
        assert sheet["E2"].value == "GOODIFSC"

        assert surat_wb["Data"]["A2"].value == 1
        assert other_wb["Data"]["A2"].value == 1
    finally:
        ahmedabad_wb.close()
        surat_wb.close()
        other_wb.close()
