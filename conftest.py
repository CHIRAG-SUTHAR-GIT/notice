"""
Pytest configuration and fixtures for the Fraud Analysis Application.

Contains hypothesis strategies for property-based testing and common fixtures.
"""

import pandas as pd
import pytest
from hypothesis import strategies as st

from src.models import (
    AggregatedAccount,
    ColumnMapping,
    ErrorCategory,
    ErrorResponse,
    ProcessingStats,
    ValidationResult,
)


# =============================================================================
# Hypothesis Strategies for Property-Based Testing
# =============================================================================

# Strategy for generating valid account numbers (9-18 digits)
valid_account_numbers = st.text(
    alphabet="0123456789",
    min_size=9,
    max_size=18
)

# Strategy for generating invalid account numbers (too short or too long)
invalid_account_numbers = st.one_of(
    st.text(alphabet="0123456789", min_size=0, max_size=8),
    st.text(alphabet="0123456789", min_size=19, max_size=25)
)

# Strategy for generating valid IFSC codes (exactly 11 alphanumeric characters)
valid_ifsc_codes = st.text(
    alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789",
    min_size=11,
    max_size=11
)

# Strategy for generating invalid IFSC codes
invalid_ifsc_codes = st.one_of(
    st.text(alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", min_size=0, max_size=10),
    st.text(alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", min_size=12, max_size=20)
)


# Strategy for generating valid transaction amounts (positive floats)
valid_amounts = st.floats(min_value=0.01, max_value=1e10, allow_nan=False, allow_infinity=False)

# Strategy for generating invalid amounts (zero or negative)
invalid_amounts = st.floats(max_value=0.0, allow_nan=False, allow_infinity=False)

# Strategy for generating bank names
bank_names = st.sampled_from([
    "State Bank of India",
    "HDFC Bank",
    "ICICI Bank",
    "Axis Bank",
    "Punjab National Bank",
    "Bank of Baroda",
    "Canara Bank",
    "Union Bank of India",
    "IndusInd Bank",
    "Kotak Mahindra Bank"
])

# Strategy for generating addresses
addresses = st.text(
    alphabet=st.characters(whitelist_categories=('L', 'N', 'P', 'Z')),
    min_size=10,
    max_size=200
)

# Strategy for generating acknowledgement numbers
acknowledgement_numbers = st.text(
    alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789",
    min_size=8,
    max_size=20
)

# Strategy for generating header strings with various formats
header_variants = st.sampled_from([
    "Bank Account No", "bank account no", "BANK ACCOUNT NO",
    "  Bank Account No  ", "bank_account_no", "Bank-Account-No",
    "Account Number", "account number", "A/C No", "ac no"
])

# Strategy for generating whitespace-padded strings
whitespace_padded_strings = st.builds(
    lambda prefix, content, suffix: prefix + content + suffix,
    prefix=st.text(alphabet=" \t", min_size=0, max_size=5),
    content=st.text(min_size=1, max_size=50),
    suffix=st.text(alphabet=" \t", min_size=0, max_size=5)
)


# Strategy for generating amount strings with currency symbols
amount_strings = st.builds(
    lambda symbol, amount, use_commas: (
        f"{symbol}{amount:,.2f}" if use_commas else f"{symbol}{amount:.2f}"
    ),
    symbol=st.sampled_from(["", "₹", "$", "Rs.", "Rs "]),
    amount=st.floats(min_value=0.01, max_value=1e8, allow_nan=False, allow_infinity=False),
    use_commas=st.booleans()
)

# Strategy for generating account numbers with spaces/dashes
account_numbers_with_formatting = st.builds(
    lambda digits, separator: separator.join(
        [digits[i:i+4] for i in range(0, len(digits), 4)]
    ),
    digits=st.text(alphabet="0123456789", min_size=9, max_size=18),
    separator=st.sampled_from([" ", "-", ""])
)


# =============================================================================
# Pytest Fixtures
# =============================================================================

@pytest.fixture
def sample_column_mapping():
    """Fixture providing a sample column mapping."""
    return ColumnMapping(
        serial_number="Sr No",
        acknowledgement_number="Ack No",
        bank_account_number="Bank Account No",
        ifsc_code="IFSC Code",
        address="Address",
        amount="Amount",
        disputed_amount="Disputed Amount",
        bank_name="Bank Name",
        confidence_scores={
            "serial_number": 1.0,
            "acknowledgement_number": 1.0,
            "bank_account_number": 1.0,
            "ifsc_code": 1.0,
            "address": 1.0,
            "amount": 1.0,
            "disputed_amount": 1.0,
            "bank_name": 1.0
        }
    )


@pytest.fixture
def sample_transaction_df():
    """Fixture providing a sample transaction DataFrame."""
    return pd.DataFrame({
        "Sr No": [1, 2, 3, 4, 5],
        "Ack No": ["ACK001", "ACK002", "ACK003", "ACK004", "ACK005"],
        "Bank Account No": [
            "123456789012",
            "123456789012",
            "987654321098",
            "987654321098",
            "555555555555"
        ],
        "IFSC Code": ["SBIN0001234", "SBIN0001234", "HDFC0005678", "HDFC0005678", "ICIC0009999"],
        "Address": [
            "123 Main St, Mumbai",
            "123 Main St, Mumbai",
            "456 Oak Ave, Delhi",
            "456 Oak Ave, Delhi",
            "789 Pine Rd, Chennai"
        ],
        "Amount": [10000.00, 15000.00, 25000.00, 30000.00, 5000.00],
        "Disputed Amount": [10000.00, 15000.00, 25000.00, 30000.00, 5000.00],
        "Bank Name": [
            "State Bank of India",
            "State Bank of India",
            "HDFC Bank",
            "HDFC Bank",
            "ICICI Bank"
        ]
    })


@pytest.fixture
def sample_aggregated_account():
    """Fixture providing a sample aggregated account."""
    return AggregatedAccount(
        account_number="123456789012",
        bank_name="State Bank of India",
        ifsc_code="SBIN0001234",
        address="123 Main St, Mumbai",
        district="Mumbai",
        state="Maharashtra",
        total_transactions=2,
        acknowledgement_numbers="ACK001;ACK002",
        total_amount=25000.00,
        total_disputed_amount=25000.00,
        risk_score=75.0
    )


@pytest.fixture
def sample_validation_result():
    """Fixture providing a sample validation result."""
    return ValidationResult(
        is_valid=True,
        critical_errors=[],
        warnings=["Duplicate acknowledgement number: ACK001"],
        flagged_rows=[],
        quality_report={
            "total_rows": 100,
            "valid_rows": 98,
            "invalid_rows": 2,
            "completeness": 0.98
        }
    )


@pytest.fixture
def empty_dataframe():
    """Fixture providing an empty DataFrame with expected columns."""
    return pd.DataFrame(columns=[
        "Sr No", "Ack No", "Bank Account No", "IFSC Code",
        "Address", "Amount", "Disputed Amount", "Bank Name"
    ])
