from types import SimpleNamespace

from src.database_service import DatabaseService
from src.models import AggregatedAccount
from src.report_generator import ReportGenerator


def test_aggregated_account_normalizes_float_text_fields():
    account = AggregatedAccount(
        account_number=1234567890.0,
        bank_name=45.5,
        ifsc_code=None,
        address=float("nan"),
        district="None",
        state="<NA>",
        total_transactions=1,
        acknowledgement_numbers=311000000000.0,
        total_amount=100.0,
        total_disputed_amount=0.0,
        risk_score=1.0,
    )

    assert account.account_number == "1234567890.0"
    assert account.bank_name == "45.5"
    assert account.ifsc_code == ""
    assert account.address == ""
    assert account.district == ""
    assert account.state == ""
    assert account.acknowledgement_numbers == "311000000000.0"


def test_report_generator_handles_legacy_float_account_fields():
    legacy_account = SimpleNamespace(
        account_number=1234567890.0,
        bank_name=45.5,
        ifsc_code=None,
        address=float("nan"),
        district="None",
        state="<NA>",
        total_transactions=1,
        acknowledgement_numbers=311000000000.0,
        total_amount=100.0,
        total_disputed_amount=0.0,
        risk_score=1.0,
    )

    df = ReportGenerator()._accounts_to_dataframe([legacy_account])

    assert df.loc[0, "Fraudster Bank Account Number"] == "1234567890.0"
    assert df.loc[0, "Bank Name"] == "45.5"
    assert df.loc[0, "Address"] == ""
    assert df.loc[0, "ACK Count"] == 1


def test_database_ack_count_handles_float_values():
    service = DatabaseService()

    assert service._calculate_ack_count(311000000000.0) == 1
    assert service._calculate_ack_count(float("nan")) == 0
