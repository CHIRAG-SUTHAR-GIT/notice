"""
Property-based tests for the Aggregation Engine.

Tests grouping uniqueness, acknowledgement consolidation, amount aggregation,
mode aggregation, transaction count, risk score determinism, and output sorting.
"""

import pandas as pd
import pytest
from hypothesis import given, settings, strategies as st, assume

from src.aggregation_engine import AggregationEngine
from src.models import AggregatedAccount, ColumnMapping


# =============================================================================
# Test Fixtures
# =============================================================================

@pytest.fixture
def aggregation_engine():
    """Fixture providing an AggregationEngine instance."""
    return AggregationEngine()


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
        bank_name="Bank Name"
    )


# =============================================================================
# Hypothesis Strategies
# =============================================================================

# Strategy for generating valid account numbers
valid_account_numbers = st.text(
    alphabet="0123456789",
    min_size=9,
    max_size=18
)

# Strategy for generating valid amounts
valid_amounts = st.floats(
    min_value=0.01, 
    max_value=1e8, 
    allow_nan=False, 
    allow_infinity=False
)

# Strategy for generating bank names
bank_names = st.sampled_from([
    "State Bank of India",
    "HDFC Bank",
    "ICICI Bank",
    "Axis Bank",
    "Punjab National Bank"
])

# Strategy for generating IFSC codes
ifsc_codes = st.text(
    alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789",
    min_size=11,
    max_size=11
)

# Strategy for generating addresses
addresses = st.text(
    alphabet=st.characters(whitelist_categories=('L', 'N', 'Z')),
    min_size=5,
    max_size=50
)

# Strategy for generating acknowledgement numbers
ack_numbers = st.text(
    alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789",
    min_size=6,
    max_size=15
)


# =============================================================================
# Property 15: Grouping Uniqueness Invariant
# =============================================================================

# Feature: fraud-analysis-app, Property 15: Grouping Uniqueness Invariant
# Validates: Requirements 4.1
# *For any* set of transactions, after aggregation by account number, the number
# of aggregated records should equal the number of unique account numbers in the input.

@settings(max_examples=100)
@given(
    account_list=st.lists(
        valid_account_numbers,
        min_size=1,
        max_size=20
    ),
    transactions_per_account=st.integers(min_value=1, max_value=5)
)
def test_property_grouping_uniqueness_invariant(account_list, transactions_per_account):
    """
    Property 15: Grouping Uniqueness Invariant
    
    For any set of transactions, after aggregation by account number, the number
    of aggregated records should equal the number of unique account numbers.
    
    Validates: Requirements 4.1
    """
    engine = AggregationEngine()
    
    # Filter out empty account numbers
    account_list = [acc for acc in account_list if acc.strip()]
    assume(len(account_list) > 0)
    
    # Create transactions with multiple entries per account
    data = []
    for account in account_list:
        for i in range(transactions_per_account):
            data.append({
                'Bank Account No': account,
                'Amount': 1000.0 + i,
                'Bank Name': 'Test Bank'
            })
    
    df = pd.DataFrame(data)
    mapping = ColumnMapping(
        bank_account_number='Bank Account No',
        amount='Amount',
        bank_name='Bank Name'
    )
    
    result = engine.aggregate_by_account(df, mapping)
    
    # Property: Number of aggregated records equals unique account count
    unique_accounts = len(set(account_list))
    assert len(result) == unique_accounts, (
        f"Expected {unique_accounts} aggregated records, got {len(result)}"
    )



# =============================================================================
# Property 16: Acknowledgement Number Consolidation
# =============================================================================

# Feature: fraud-analysis-app, Property 16: Acknowledgement Number Consolidation
# Validates: Requirements 4.2
# *For any* group of transactions with the same account number, the consolidated
# acknowledgement numbers string should contain all individual acknowledgement
# numbers from that group, separated by semicolons.

@settings(max_examples=100)
@given(
    account_number=valid_account_numbers,
    ack_list=st.lists(ack_numbers, min_size=1, max_size=10)
)
def test_property_acknowledgement_consolidation(account_number, ack_list):
    """
    Property 16: Acknowledgement Number Consolidation
    
    For any group of transactions with the same account number, the consolidated
    acknowledgement numbers string should contain all individual acknowledgement
    numbers from that group, separated by semicolons.
    
    Validates: Requirements 4.2
    """
    engine = AggregationEngine()
    
    # Filter out empty ack numbers
    ack_list = [ack for ack in ack_list if ack.strip()]
    assume(len(ack_list) > 0)
    assume(account_number.strip())
    
    # Create transactions with the same account but different ack numbers
    data = []
    for ack in ack_list:
        data.append({
            'Bank Account No': account_number,
            'Ack No': ack,
            'Amount': 1000.0
        })
    
    df = pd.DataFrame(data)
    mapping = ColumnMapping(
        bank_account_number='Bank Account No',
        acknowledgement_number='Ack No',
        amount='Amount'
    )
    
    result = engine.aggregate_by_account(df, mapping)
    
    assert len(result) == 1, "Should have exactly one aggregated account"
    
    # Property: All ack numbers should be in the consolidated string
    consolidated = result[0].acknowledgement_numbers
    consolidated_set = set(consolidated.split(';'))
    
    for ack in ack_list:
        assert ack in consolidated_set, (
            f"Acknowledgement number '{ack}' not found in consolidated string '{consolidated}'"
        )


def test_acknowledgement_deduplication():
    """
    Test that duplicate acknowledgement numbers are properly removed.
    
    This test specifically addresses the issue where the same acknowledgement
    number appears multiple times in the consolidated string.
    """
    engine = AggregationEngine()
    
    # Create test data with duplicate acknowledgement numbers
    df = pd.DataFrame({
        'Bank Account No': ['111111111', '111111111', '111111111', '111111111'],
        'Ack No': ['ACK001', 'ACK001', 'ACK002', 'ACK001'],  # ACK001 appears 3 times
        'Amount': [1000.0, 2000.0, 3000.0, 4000.0]
    })
    
    mapping = ColumnMapping(
        bank_account_number='Bank Account No',
        acknowledgement_number='Ack No',
        amount='Amount'
    )
    
    result = engine.aggregate_by_account(df, mapping)
    
    assert len(result) == 1, "Should have exactly one aggregated account"
    
    # Check that acknowledgement numbers are deduplicated
    consolidated = result[0].acknowledgement_numbers
    ack_numbers = consolidated.split(';')
    
    # Should only have 2 unique acknowledgement numbers
    unique_acks = set(ack_numbers)
    assert len(unique_acks) == 2, f"Expected 2 unique ack numbers, got {len(unique_acks)}: {unique_acks}"
    assert 'ACK001' in unique_acks
    assert 'ACK002' in unique_acks
    
    # Should not have duplicates in the consolidated string
    assert len(ack_numbers) == len(unique_acks), f"Found duplicates in consolidated string: {ack_numbers}"


# =============================================================================
# Property 17: Amount Sum Aggregation
# =============================================================================

# Feature: fraud-analysis-app, Property 17: Amount Sum Aggregation
# Validates: Requirements 4.3, 4.4
# *For any* group of transactions with the same account number, the total_amount
# in the aggregated record should equal the sum of all individual transaction amounts.

@settings(max_examples=100)
@given(
    account_number=valid_account_numbers,
    amounts=st.lists(valid_amounts, min_size=1, max_size=10),
    disputed_amounts=st.lists(valid_amounts, min_size=1, max_size=10)
)
def test_property_amount_sum_aggregation(account_number, amounts, disputed_amounts):
    """
    Property 17: Amount Sum Aggregation
    
    For any group of transactions with the same account number, the total_amount
    should equal the sum of all individual transaction amounts. Similarly for
    total_disputed_amount.
    
    Validates: Requirements 4.3, 4.4
    """
    engine = AggregationEngine()
    
    assume(account_number.strip())
    
    # Ensure lists are same length
    min_len = min(len(amounts), len(disputed_amounts))
    amounts = amounts[:min_len]
    disputed_amounts = disputed_amounts[:min_len]
    
    # Create transactions
    data = []
    for amt, disp in zip(amounts, disputed_amounts):
        data.append({
            'Bank Account No': account_number,
            'Amount': amt,
            'Disputed Amount': disp
        })
    
    df = pd.DataFrame(data)
    mapping = ColumnMapping(
        bank_account_number='Bank Account No',
        amount='Amount',
        disputed_amount='Disputed Amount'
    )
    
    result = engine.aggregate_by_account(df, mapping)
    
    assert len(result) == 1, "Should have exactly one aggregated account"
    
    # Property: Total amount should equal sum of individual amounts
    expected_total = sum(amounts)
    assert abs(result[0].total_amount - expected_total) < 0.01, (
        f"Total amount {result[0].total_amount} does not match expected {expected_total}"
    )
    
    # Property: Total disputed amount should equal sum of individual disputed amounts
    expected_disputed = sum(disputed_amounts)
    assert abs(result[0].total_disputed_amount - expected_disputed) < 0.01, (
        f"Total disputed {result[0].total_disputed_amount} does not match expected {expected_disputed}"
    )



# =============================================================================
# Property 18: Mode Aggregation for Categorical Fields
# =============================================================================

# Feature: fraud-analysis-app, Property 18: Mode Aggregation for Categorical Fields
# Validates: Requirements 4.5, 4.6, 4.7
# *For any* group of transactions with the same account number, the bank_name,
# ifsc_code, and address should each be the most frequently occurring (mode)
# non-null value from that group.

@settings(max_examples=100)
@given(
    account_number=valid_account_numbers,
    dominant_bank=bank_names,
    other_bank=bank_names,
    dominant_count=st.integers(min_value=3, max_value=10),
    other_count=st.integers(min_value=1, max_value=2)
)
def test_property_mode_aggregation_categorical(
    account_number, dominant_bank, other_bank, dominant_count, other_count
):
    """
    Property 18: Mode Aggregation for Categorical Fields
    
    For any group of transactions with the same account number, the bank_name,
    ifsc_code, and address should each be the most frequently occurring (mode)
    non-null value from that group.
    
    Validates: Requirements 4.5, 4.6, 4.7
    """
    engine = AggregationEngine()
    
    assume(account_number.strip())
    assume(dominant_bank != other_bank)  # Ensure different banks
    
    # Create transactions with dominant bank appearing more often
    data = []
    for _ in range(dominant_count):
        data.append({
            'Bank Account No': account_number,
            'Bank Name': dominant_bank,
            'Amount': 1000.0
        })
    for _ in range(other_count):
        data.append({
            'Bank Account No': account_number,
            'Bank Name': other_bank,
            'Amount': 1000.0
        })
    
    df = pd.DataFrame(data)
    mapping = ColumnMapping(
        bank_account_number='Bank Account No',
        bank_name='Bank Name',
        amount='Amount'
    )
    
    result = engine.aggregate_by_account(df, mapping)
    
    assert len(result) == 1, "Should have exactly one aggregated account"
    
    # Property: Bank name should be the most common (dominant) one
    assert result[0].bank_name == dominant_bank, (
        f"Expected bank name '{dominant_bank}', got '{result[0].bank_name}'"
    )


# =============================================================================
# Property 19: Transaction Count Accuracy
# =============================================================================

# Feature: fraud-analysis-app, Property 19: Transaction Count Accuracy
# Validates: Requirements 4.8
# *For any* group of transactions with the same account number, the
# total_transactions count should equal the number of rows in that group.

@settings(max_examples=100)
@given(
    account_number=valid_account_numbers,
    num_transactions=st.integers(min_value=1, max_value=50)
)
def test_property_transaction_count_accuracy(account_number, num_transactions):
    """
    Property 19: Transaction Count Accuracy
    
    For any group of transactions with the same account number, the
    total_transactions count should equal the number of rows in that group.
    
    Validates: Requirements 4.8
    """
    engine = AggregationEngine()
    
    assume(account_number.strip())
    
    # Create transactions
    data = []
    for i in range(num_transactions):
        data.append({
            'Bank Account No': account_number,
            'Amount': 1000.0 + i
        })
    
    df = pd.DataFrame(data)
    mapping = ColumnMapping(
        bank_account_number='Bank Account No',
        amount='Amount'
    )
    
    result = engine.aggregate_by_account(df, mapping)
    
    assert len(result) == 1, "Should have exactly one aggregated account"
    
    # Property: Transaction count should equal number of rows
    assert result[0].total_transactions == num_transactions, (
        f"Expected {num_transactions} transactions, got {result[0].total_transactions}"
    )



# =============================================================================
# Property 20: Risk Score Determinism
# =============================================================================

# Feature: fraud-analysis-app, Property 20: Risk Score Determinism
# Validates: Requirements 4.9
# *For any* two aggregated accounts with identical transaction counts and total
# amounts, their calculated risk scores should be equal.

@settings(max_examples=100)
@given(
    transaction_count=st.integers(min_value=1, max_value=200),
    total_amount=st.floats(min_value=0.01, max_value=1e9, allow_nan=False, allow_infinity=False)
)
def test_property_risk_score_determinism(transaction_count, total_amount):
    """
    Property 20: Risk Score Determinism
    
    For any two aggregated accounts with identical transaction counts and total
    amounts, their calculated risk scores should be equal.
    
    Validates: Requirements 4.9
    """
    engine = AggregationEngine()
    
    # Calculate risk score twice with same inputs
    score1 = engine.calculate_risk_score(transaction_count, total_amount)
    score2 = engine.calculate_risk_score(transaction_count, total_amount)
    
    # Property: Same inputs should produce same output
    assert score1 == score2, (
        f"Risk scores differ for same inputs: {score1} vs {score2}"
    )
    
    # Property: Risk score should be between 0 and 100
    assert 0 <= score1 <= 100, (
        f"Risk score {score1} is outside valid range [0, 100]"
    )


# =============================================================================
# Property 21: Output Sorting Order
# =============================================================================

# Feature: fraud-analysis-app, Property 21: Output Sorting Order
# Validates: Requirements 5.2, 5.3
# *For any* list of aggregated accounts in the output, accounts should be sorted
# such that for any two adjacent accounts A and B (where A comes before B):
# either A.total_amount > B.total_amount, or (A.total_amount == B.total_amount
# and A.total_transactions >= B.total_transactions).

@settings(max_examples=100)
@given(
    accounts_data=st.lists(
        st.tuples(
            valid_account_numbers,
            valid_amounts,
            st.integers(min_value=1, max_value=100)
        ),
        min_size=2,
        max_size=20
    )
)
def test_property_output_sorting_order(accounts_data):
    """
    Property 21: Output Sorting Order
    
    For any list of aggregated accounts, after sorting, accounts should be
    ordered by total_amount descending, then by total_transactions descending.
    
    Validates: Requirements 5.2, 5.3
    """
    engine = AggregationEngine()
    
    # Filter out empty account numbers
    accounts_data = [(acc, amt, txn) for acc, amt, txn in accounts_data if acc.strip()]
    assume(len(accounts_data) >= 2)
    
    # Create AggregatedAccount objects
    accounts = []
    for acc, amt, txn in accounts_data:
        accounts.append(AggregatedAccount(
            account_number=acc,
            bank_name="Test Bank",
            ifsc_code="TEST0001234",
            address="Test Address",
            total_transactions=txn,
            acknowledgement_numbers="ACK001",
            total_amount=amt,
            total_disputed_amount=0.0,
            risk_score=50.0
        ))
    
    # Sort the accounts
    sorted_accounts = engine.sort_results(accounts)
    
    # Property: Verify sorting order
    for i in range(len(sorted_accounts) - 1):
        current = sorted_accounts[i]
        next_acc = sorted_accounts[i + 1]
        
        # Either current amount > next amount
        # OR amounts are equal and current transactions >= next transactions
        valid_order = (
            current.total_amount > next_acc.total_amount or
            (current.total_amount == next_acc.total_amount and 
             current.total_transactions >= next_acc.total_transactions)
        )
        
        assert valid_order, (
            f"Invalid sort order: account with amount {current.total_amount} "
            f"and {current.total_transactions} transactions comes before "
            f"account with amount {next_acc.total_amount} and "
            f"{next_acc.total_transactions} transactions"
        )



# =============================================================================
# Unit Tests for AggregationEngine
# =============================================================================

class TestGetMostCommon:
    """Unit tests for get_most_common method."""
    
    def test_returns_most_frequent_value(self, aggregation_engine):
        """Test that the most frequent value is returned."""
        series = pd.Series(['A', 'B', 'A', 'A', 'B'])
        result = aggregation_engine.get_most_common(series)
        assert result == 'A'
    
    def test_handles_all_same_values(self, aggregation_engine):
        """Test handling when all values are the same."""
        series = pd.Series(['X', 'X', 'X'])
        result = aggregation_engine.get_most_common(series)
        assert result == 'X'
    
    def test_handles_empty_series(self, aggregation_engine):
        """Test handling of empty series."""
        series = pd.Series([], dtype=object)
        result = aggregation_engine.get_most_common(series)
        assert result == ''
    
    def test_handles_all_null_values(self, aggregation_engine):
        """Test handling when all values are null."""
        series = pd.Series([None, None, None])
        result = aggregation_engine.get_most_common(series)
        assert result == ''
    
    def test_ignores_null_values(self, aggregation_engine):
        """Test that null values are ignored in mode calculation."""
        series = pd.Series(['A', None, 'B', None, 'A'])
        result = aggregation_engine.get_most_common(series)
        assert result == 'A'
    
    def test_ignores_empty_strings(self, aggregation_engine):
        """Test that empty strings are ignored."""
        series = pd.Series(['A', '', 'B', '', 'A', ''])
        result = aggregation_engine.get_most_common(series)
        assert result == 'A'


class TestCalculateRiskScore:
    """Unit tests for calculate_risk_score method."""
    
    def test_low_risk_score(self, aggregation_engine):
        """Test low risk score for few transactions and low amount."""
        score = aggregation_engine.calculate_risk_score(1, 1000.0)
        assert score < 10
    
    def test_high_risk_score(self, aggregation_engine):
        """Test high risk score for many transactions and high amount."""
        score = aggregation_engine.calculate_risk_score(100, 10_000_000.0)
        assert score >= 90
    
    def test_score_in_valid_range(self, aggregation_engine):
        """Test that score is always between 0 and 100."""
        score = aggregation_engine.calculate_risk_score(50, 5_000_000.0)
        assert 0 <= score <= 100
    
    def test_zero_transactions(self, aggregation_engine):
        """Test handling of zero transactions."""
        score = aggregation_engine.calculate_risk_score(0, 1000.0)
        assert score >= 0


class TestAggregateByAccount:
    """Unit tests for aggregate_by_account method."""
    
    def test_aggregates_multiple_accounts(self, aggregation_engine, sample_column_mapping):
        """Test aggregation of multiple accounts."""
        df = pd.DataFrame({
            'Sr No': [1, 2, 3, 4],
            'Ack No': ['ACK001', 'ACK002', 'ACK003', 'ACK004'],
            'Bank Account No': ['111111111', '111111111', '222222222', '222222222'],
            'IFSC Code': ['SBIN0001234', 'SBIN0001234', 'HDFC0005678', 'HDFC0005678'],
            'Address': ['Addr1', 'Addr1', 'Addr2', 'Addr2'],
            'Amount': [1000.0, 2000.0, 3000.0, 4000.0],
            'Disputed Amount': [500.0, 1000.0, 1500.0, 2000.0],
            'Bank Name': ['SBI', 'SBI', 'HDFC', 'HDFC']
        })
        
        result = aggregation_engine.aggregate_by_account(df, sample_column_mapping)
        
        assert len(result) == 2
        
        # Find account 111111111
        acc1 = next(a for a in result if a.account_number == '111111111')
        assert acc1.total_amount == 3000.0
        assert acc1.total_disputed_amount == 1500.0
        assert acc1.total_transactions == 2
        assert 'ACK001' in acc1.acknowledgement_numbers
        assert 'ACK002' in acc1.acknowledgement_numbers
    
    def test_handles_empty_dataframe(self, aggregation_engine, sample_column_mapping):
        """Test handling of empty DataFrame."""
        df = pd.DataFrame(columns=[
            'Sr No', 'Ack No', 'Bank Account No', 'IFSC Code',
            'Address', 'Amount', 'Disputed Amount', 'Bank Name'
        ])
        
        result = aggregation_engine.aggregate_by_account(df, sample_column_mapping)
        
        assert len(result) == 0
    
    def test_handles_missing_optional_columns(self, aggregation_engine):
        """Test handling when optional columns are missing."""
        df = pd.DataFrame({
            'Bank Account No': ['111111111', '111111111'],
            'Amount': [1000.0, 2000.0]
        })
        
        mapping = ColumnMapping(
            bank_account_number='Bank Account No',
            amount='Amount'
        )
        
        result = aggregation_engine.aggregate_by_account(df, mapping)
        
        assert len(result) == 1
        assert result[0].total_amount == 3000.0
        assert result[0].bank_name == ''
        assert result[0].ifsc_code == ''
    
    def test_skips_null_account_numbers(self, aggregation_engine, sample_column_mapping):
        """Test that rows with null account numbers are skipped."""
        df = pd.DataFrame({
            'Sr No': [1, 2, 3],
            'Ack No': ['ACK001', 'ACK002', 'ACK003'],
            'Bank Account No': ['111111111', None, '111111111'],
            'IFSC Code': ['SBIN0001234', 'SBIN0001234', 'SBIN0001234'],
            'Address': ['Addr1', 'Addr1', 'Addr1'],
            'Amount': [1000.0, 2000.0, 3000.0],
            'Disputed Amount': [500.0, 1000.0, 1500.0],
            'Bank Name': ['SBI', 'SBI', 'SBI']
        })
        
        result = aggregation_engine.aggregate_by_account(df, sample_column_mapping)
        
        assert len(result) == 1
        assert result[0].total_amount == 4000.0  # Only rows 1 and 3


class TestSortResults:
    """Unit tests for sort_results method."""
    
    def test_sorts_by_amount_descending(self, aggregation_engine):
        """Test that accounts are sorted by amount descending."""
        accounts = [
            AggregatedAccount('111', 'Bank1', 'IFSC1', 'Addr1', "", "", 1, 'ACK1', 1000.0, 0.0, 50.0),
            AggregatedAccount('222', 'Bank2', 'IFSC2', 'Addr2', "", "", 1, 'ACK2', 3000.0, 0.0, 50.0),
            AggregatedAccount('333', 'Bank3', 'IFSC3', 'Addr3', "", "", 1, 'ACK3', 2000.0, 0.0, 50.0),
        ]
        
        result = aggregation_engine.sort_results(accounts)
        
        assert result[0].total_amount == 3000.0
        assert result[1].total_amount == 2000.0
        assert result[2].total_amount == 1000.0
    
    def test_sorts_by_transactions_when_amounts_equal(self, aggregation_engine):
        """Test secondary sort by transaction count when amounts are equal."""
        accounts = [
            AggregatedAccount('111', 'Bank1', 'IFSC1', 'Addr1', "", "", 5, 'ACK1', 1000.0, 0.0, 50.0),
            AggregatedAccount('222', 'Bank2', 'IFSC2', 'Addr2', "", "", 10, 'ACK2', 1000.0, 0.0, 50.0),
            AggregatedAccount('333', 'Bank3', 'IFSC3', 'Addr3', "", "", 3, 'ACK3', 1000.0, 0.0, 50.0),
        ]
        
        result = aggregation_engine.sort_results(accounts)
        
        assert result[0].total_transactions == 10
        assert result[1].total_transactions == 5
        assert result[2].total_transactions == 3
    
    def test_handles_empty_list(self, aggregation_engine):
        """Test handling of empty list."""
        result = aggregation_engine.sort_results([])
        assert result == []
    
    def test_handles_single_account(self, aggregation_engine):
        """Test handling of single account."""
        accounts = [
            AggregatedAccount('111', 'Bank1', 'IFSC1', 'Addr1', "", "", 1, 'ACK1', 1000.0, 0.0, 50.0)
        ]
        
        result = aggregation_engine.sort_results(accounts)
        
        assert len(result) == 1
        assert result[0].account_number == '111'
