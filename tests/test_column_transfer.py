import json

import pandas as pd

from src.column_transfer import load_settings, normalize_match_value, save_settings, transfer_columns


def test_transfer_columns_inserts_renamed_file2_columns_after_selected_base_column():
    base_df = pd.DataFrame(
        {
            "Sr No": [1, 2, 3],
            "Ack No": ["311-001", "311-002", "311-003"],
            "Amount": [100, 200, 300],
        }
    )
    lookup_df = pd.DataFrame(
        {
            "Acknowledgement No.": ["311001", "311002"],
            "Bank/FIs": ["SBI", "HDFC"],
            "Account No.": ["1001", "2002"],
        }
    )

    output, stats, final_names = transfer_columns(
        base_df=base_df,
        lookup_df=lookup_df,
        base_match_col="Ack No",
        lookup_match_col="Acknowledgement No.",
        selected_lookup_cols=["Bank/FIs", "Account No."],
        rename_map={"Bank/FIs": "Bank Name", "Account No.": "AC NO"},
        insert_after="Ack No",
        match_mode="smart",
    )

    assert list(output.columns) == ["Sr No", "Ack No", "Bank Name", "AC NO", "Amount"]
    assert output["Bank Name"].tolist() == ["SBI", "HDFC", ""]
    assert output["AC NO"].tolist() == ["1001", "2002", ""]
    assert stats["matched_rows"] == 2
    assert stats["unmatched_rows"] == 1
    assert final_names == {"Bank/FIs": "Bank Name", "Account No.": "AC NO"}


def test_transfer_columns_uses_first_lookup_duplicate_key_without_multiplying_rows():
    base_df = pd.DataFrame({"Account": ["1001", "2002"]})
    lookup_df = pd.DataFrame(
        {
            "Account No": ["1001", "1001", "2002"],
            "Mobile": ["900001", "900002", "800001"],
        }
    )

    output, stats, _ = transfer_columns(
        base_df,
        lookup_df,
        "Account",
        "Account No",
        ["Mobile"],
    )

    assert len(output) == 2
    assert output["Mobile"].tolist() == ["900001", "800001"]
    assert stats["duplicate_lookup_keys"] == 1


def test_transfer_columns_handles_file2_column_name_already_in_file1():
    base_df = pd.DataFrame({"Account": ["1001"], "Bank": ["OLD"]})
    lookup_df = pd.DataFrame({"Account": ["1001"], "Bank": ["NEW"]})

    output, _, final_names = transfer_columns(
        base_df,
        lookup_df,
        "Account",
        "Account",
        ["Bank"],
        rename_map={"Bank": "Bank"},
    )

    assert output["Bank"].tolist() == ["OLD"]
    assert output["Bank_2"].tolist() == ["NEW"]
    assert final_names["Bank"] == "Bank_2"


def test_transfer_columns_can_rename_base_file_columns():
    base_df = pd.DataFrame({"Ack": ["311001"], "Old Amount": ["100"]})
    lookup_df = pd.DataFrame({"Ack": ["311001"], "District": ["Rajkot"]})

    output, stats, final_names = transfer_columns(
        base_df,
        lookup_df,
        "Ack",
        "Ack",
        ["District"],
        base_rename_map={"Ack": "ACK NO", "Old Amount": "Amount Reported"},
        insert_after="Ack",
    )

    assert list(output.columns) == ["ACK NO", "District", "Amount Reported"]
    assert output.loc[0, "ACK NO"] == "311001"
    assert stats["base_columns_renamed"] == 2
    assert final_names["District"] == "District"


def test_base_rename_conflict_with_added_column_gets_unique_name():
    base_df = pd.DataFrame({"Ack": ["311001"], "Bank": ["OLD"]})
    lookup_df = pd.DataFrame({"Ack": ["311001"], "District": ["Rajkot"]})

    output, _, final_names = transfer_columns(
        base_df,
        lookup_df,
        "Ack",
        "Ack",
        ["District"],
        rename_map={"District": "Bank"},
        base_rename_map={"Bank": "Bank"},
        insert_after="Ack",
    )

    assert list(output.columns) == ["Ack", "Bank_2", "Bank"]
    assert output.loc[0, "Bank"] == "OLD"
    assert output.loc[0, "Bank_2"] == "Rajkot"
    assert final_names["District"] == "Bank_2"


def test_match_modes():
    assert normalize_match_value("311-001.0", "smart") == "311001"
    assert normalize_match_value(" A B 12 ", "exact") == "A B 12"


def test_settings_round_trip(tmp_path):
    path = tmp_path / "settings.json"
    settings = {
        "base_match_col": "Ack No",
        "lookup_match_col": "Ack",
        "selected_lookup_cols": ["Bank"],
        "rename_map": {"Bank": "Bank Name"},
        "base_rename_map": {"Ack No": "ACK NO"},
        "insert_after": "Ack No",
    }

    save_settings(settings, path)

    assert json.loads(path.read_text(encoding="utf-8")) == settings
    assert load_settings(path) == settings
