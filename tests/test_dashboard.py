"""
Property-based tests for the Dashboard module.

Tests statistics calculation consistency, search filter correctness,
and minimum filter correctness.
"""

import pytest
from hypothesis import given, settings, strategies as st, assume

from src.dashboard import Dashboard
from src.models import AggregatedAccount, ProcessingStats


# =============================================================================
# Test Fixtures
# =============================================================================

@pytest.fixture
def dashboard():
    """Fixture providing a Dashboard instance."""
    return Dashboard()


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


# Strategy for generating AggregatedAccount objects
@st.composite
def aggregated_account_strategy(draw):
    """Strategy for generating valid AggregatedAccount objects."""
    account_number = draw(valid_account_numbers)
    assume(account_number.strip())
    
    return AggregatedAccount(
        account_number=account_number,
        bank_name=draw(bank_names),
        ifsc_code=draw(ifsc_codes),
        address=draw(addresses),
        district=draw(st.text(min_size=0, max_size=30)),
        state=draw(st.text(min_size=0, max_size=30)),
        total_transactions=draw(st.integers(min_value=1, max_value=100)),
        acknowledgement_numbers=draw(ack_numbers),
        total_amount=draw(valid_amounts),
        total_disputed_amount=draw(valid_amounts),
        risk_score=draw(st.floats(min_value=0.0, max_value=100.0, allow_nan=False))
    )


# =============================================================================
# Property 24: Statistics Calculation Consistency
# =============================================================================

# Feature: fraud-analysis-app, Property 24: Statistics Calculation Consistency
# Validates: Requirements 6.1, 6.2, 6.3
# *For any* processing result, the total_fraud_amount statistic should equal
# the sum of total_amount across all aggregated accounts, and unique_accounts
# should equal the count of aggregated records.

@settings(max_examples=100)
@given(
    accounts=st.lists(aggregated_account_strategy(), min_size=0, max_size=20),
    total_input_rows=st.integers(min_value=1, max_value=1000),
    rows_with_errors=st.integers(min_value=0, max_value=100)
)
def test_property_statistics_calculation_consistency(accounts, total_input_rows, rows_with_errors):
    """
    Property 24: Statistics Calculation Consistency
    
    For any processing result, the total_fraud_amount statistic should equal
    the sum of total_amount across all aggregated accounts, and unique_accounts
    should equal the count of aggregated records.
    
    Validates: Requirements 6.1, 6.2, 6.3
    """
    dashboard = Dashboard()
    
    # Ensure rows_with_errors doesn't exceed total_input_rows
    rows_with_errors = min(rows_with_errors, total_input_rows)
    
    stats = dashboard.calculate_statistics(
        accounts=accounts,
        total_input_rows=total_input_rows,
        input_filename="test.xlsx",
        rows_with_errors=rows_with_errors
    )
    
    # Property: unique_accounts should equal count of aggregated records
    assert stats.unique_accounts == len(accounts), (
        f"Expected unique_accounts={len(accounts)}, got {stats.unique_accounts}"
    )
    
    # Property: total_fraud_amount should equal sum of all total_amounts
    expected_total = sum(acc.total_amount for acc in accounts)
    assert abs(stats.total_fraud_amount - expected_total) < 0.01, (
        f"Expected total_fraud_amount={expected_total}, got {stats.total_fraud_amount}"
    )
    
    # Property: total_disputed_amount should equal sum of all disputed amounts
    expected_disputed = sum(acc.total_disputed_amount for acc in accounts)
    assert abs(stats.total_disputed_amount - expected_disputed) < 0.01, (
        f"Expected total_disputed_amount={expected_disputed}, got {stats.total_disputed_amount}"
    )
    
    # Property: average_amount_per_account should be correct
    if len(accounts) > 0:
        expected_avg = expected_total / len(accounts)
        assert abs(stats.average_amount_per_account - expected_avg) < 0.01, (
            f"Expected average={expected_avg}, got {stats.average_amount_per_account}"
        )
    else:
        assert stats.average_amount_per_account == 0.0


# =============================================================================
# Property 25: Search Filter Correctness
# =============================================================================

# Feature: fraud-analysis-app, Property 25: Search Filter Correctness
# Validates: Requirements 6.6
# *For any* search query for an account number, the filtered results should
# contain only accounts where the account number matches or contains the search query.

@settings(max_examples=100)
@given(
    accounts=st.lists(aggregated_account_strategy(), min_size=1, max_size=20),
    query_index=st.integers(min_value=0, max_value=100)
)
def test_property_search_filter_correctness(accounts, query_index):
    """
    Property 25: Search Filter Correctness
    
    For any search query for an account number, the filtered results should
    contain only accounts where the account number matches or contains the search query.
    
    Validates: Requirements 6.6
    """
    dashboard = Dashboard()
    
    assume(len(accounts) > 0)
    
    # Pick a substring from one of the account numbers as the query
    target_account = accounts[query_index % len(accounts)]
    account_num = target_account.account_number
    
    # Use a substring of the account number as query
    if len(account_num) >= 3:
        query = account_num[:3]
    else:
        query = account_num
    
    results = dashboard.search_accounts(accounts, query)
    
    # Property: All results should contain the query in their account number
    for acc in results:
        assert query in acc.account_number, (
            f"Account {acc.account_number} does not contain query '{query}'"
        )
    
    # Property: All accounts containing the query should be in results
    expected_accounts = [acc for acc in accounts if query in acc.account_number]
    assert len(results) == len(expected_accounts), (
        f"Expected {len(expected_accounts)} results, got {len(results)}"
    )


@settings(max_examples=100)
@given(
    accounts=st.lists(aggregated_account_strategy(), min_size=1, max_size=20)
)
def test_property_empty_search_returns_all(accounts):
    """
    Property 25 (edge case): Empty search query should return all accounts.
    
    Validates: Requirements 6.6
    """
    dashboard = Dashboard()
    
    # Empty query should return all accounts
    results = dashboard.search_accounts(accounts, "")
    assert len(results) == len(accounts)
    
    # Whitespace-only query should also return all accounts
    results = dashboard.search_accounts(accounts, "   ")
    assert len(results) == len(accounts)


# =============================================================================
# Property 26: Minimum Filter Correctness
# =============================================================================

# Feature: fraud-analysis-app, Property 26: Minimum Filter Correctness
# Validates: Requirements 6.7, 6.8
# *For any* minimum transaction count filter N, the filtered results should
# contain only accounts where total_transactions >= N. Similarly for minimum amount filter.

@settings(max_examples=100)
@given(
    accounts=st.lists(aggregated_account_strategy(), min_size=1, max_size=20),
    min_transactions=st.integers(min_value=1, max_value=50)
)
def test_property_min_transactions_filter_correctness(accounts, min_transactions):
    """
    Property 26: Minimum Filter Correctness (Transactions)
    
    For any minimum transaction count filter N, the filtered results should
    contain only accounts where total_transactions >= N.
    
    Validates: Requirements 6.7
    """
    dashboard = Dashboard()
    
    results = dashboard.filter_by_min_transactions(accounts, min_transactions)
    
    # Property: All results should have transactions >= min_transactions
    for acc in results:
        assert acc.total_transactions >= min_transactions, (
            f"Account has {acc.total_transactions} transactions, "
            f"expected >= {min_transactions}"
        )
    
    # Property: All accounts meeting criteria should be in results
    expected_accounts = [acc for acc in accounts if acc.total_transactions >= min_transactions]
    assert len(results) == len(expected_accounts), (
        f"Expected {len(expected_accounts)} results, got {len(results)}"
    )


@settings(max_examples=100)
@given(
    accounts=st.lists(aggregated_account_strategy(), min_size=1, max_size=20),
    min_amount=st.floats(min_value=0.01, max_value=1e7, allow_nan=False, allow_infinity=False)
)
def test_property_min_amount_filter_correctness(accounts, min_amount):
    """
    Property 26: Minimum Filter Correctness (Amount)
    
    For any minimum amount filter M, the filtered results should
    contain only accounts where total_amount >= M.
    
    Validates: Requirements 6.8
    """
    dashboard = Dashboard()
    
    results = dashboard.filter_by_min_amount(accounts, min_amount)
    
    # Property: All results should have amount >= min_amount
    for acc in results:
        assert acc.total_amount >= min_amount, (
            f"Account has amount {acc.total_amount}, expected >= {min_amount}"
        )
    
    # Property: All accounts meeting criteria should be in results
    expected_accounts = [acc for acc in accounts if acc.total_amount >= min_amount]
    assert len(results) == len(expected_accounts), (
        f"Expected {len(expected_accounts)} results, got {len(results)}"
    )


@settings(max_examples=100)
@given(
    accounts=st.lists(aggregated_account_strategy(), min_size=1, max_size=20)
)
def test_property_zero_filter_returns_all(accounts):
    """
    Property 26 (edge case): Zero or negative filter values should return all accounts.
    
    Validates: Requirements 6.7, 6.8
    """
    dashboard = Dashboard()
    
    # Zero min_transactions should return all
    results = dashboard.filter_by_min_transactions(accounts, 0)
    assert len(results) == len(accounts)
    
    # Negative min_transactions should return all
    results = dashboard.filter_by_min_transactions(accounts, -5)
    assert len(results) == len(accounts)
    
    # Zero min_amount should return all
    results = dashboard.filter_by_min_amount(accounts, 0.0)
    assert len(results) == len(accounts)
    
    # Negative min_amount should return all
    results = dashboard.filter_by_min_amount(accounts, -100.0)
    assert len(results) == len(accounts)


# =============================================================================
# Unit Tests for Dashboard
# =============================================================================

class TestCalculateStatistics:
    """Unit tests for calculate_statistics method."""
    
    def test_empty_accounts_list(self, dashboard):
        """Test statistics calculation with empty accounts list."""
        stats = dashboard.calculate_statistics(
            accounts=[],
            total_input_rows=100,
            input_filename="test.xlsx",
            rows_with_errors=5
        )
        
        assert stats.unique_accounts == 0
        assert stats.total_fraud_amount == 0.0
        assert stats.total_disputed_amount == 0.0
        assert stats.average_amount_per_account == 0.0
        assert stats.top_accounts_by_amount == []
        assert stats.total_input_rows == 100
        assert stats.rows_with_errors == 5
    
    def test_top_accounts_limited_to_20(self, dashboard):
        """Test that top_accounts_by_amount is limited to 20."""
        accounts = [
            AggregatedAccount(
                account_number=f"12345678901{i}",
                bank_name="Test Bank",
                ifsc_code="TEST0001234",
                address="Test Address",
                district="Test District",
                state="Test State",
                total_transactions=i,
                acknowledgement_numbers=f"ACK{i}",
                total_amount=float(i * 1000),
                total_disputed_amount=0.0,
                risk_score=50.0
            )
            for i in range(1, 26)  # 25 accounts
        ]
        
        stats = dashboard.calculate_statistics(
            accounts=accounts,
            total_input_rows=100,
            input_filename="test.xlsx"
        )
        
        assert len(stats.top_accounts_by_amount) == 20
        # Verify they are sorted by amount descending
        for i in range(len(stats.top_accounts_by_amount) - 1):
            assert stats.top_accounts_by_amount[i].total_amount >= \
                   stats.top_accounts_by_amount[i + 1].total_amount


class TestSearchAccounts:
    """Unit tests for search_accounts method."""
    
    def test_exact_match(self, dashboard):
        """Test search with exact account number match."""
        accounts = [
            AggregatedAccount("123456789012", "Bank1", "IFSC1", "Addr1", "", "", 1, "ACK1", 1000.0, 0.0, 50.0),
            AggregatedAccount("987654321098", "Bank2", "IFSC2", "Addr2", "", "", 1, "ACK2", 2000.0, 0.0, 50.0),
        ]
        
        results = dashboard.search_accounts(accounts, "123456789012")
        
        assert len(results) == 1
        assert results[0].account_number == "123456789012"
    
    def test_partial_match(self, dashboard):
        """Test search with partial account number match."""
        accounts = [
            AggregatedAccount("123456789012", "Bank1", "IFSC1", "Addr1", "", "", 1, "ACK1", 1000.0, 0.0, 50.0),
            AggregatedAccount("123456000000", "Bank2", "IFSC2", "Addr2", "", "", 1, "ACK2", 2000.0, 0.0, 50.0),
            AggregatedAccount("987654321098", "Bank3", "IFSC3", "Addr3", "", "", 1, "ACK3", 3000.0, 0.0, 50.0),
        ]
        
        results = dashboard.search_accounts(accounts, "123456")
        
        assert len(results) == 2
    
    def test_no_match(self, dashboard):
        """Test search with no matching accounts."""
        accounts = [
            AggregatedAccount("123456789012", "Bank1", "IFSC1", "Addr1", "", "", 1, "ACK1", 1000.0, 0.0, 50.0),
        ]
        
        results = dashboard.search_accounts(accounts, "999999")
        
        assert len(results) == 0


class TestFilterByMinTransactions:
    """Unit tests for filter_by_min_transactions method."""
    
    def test_filters_correctly(self, dashboard):
        """Test filtering by minimum transactions."""
        accounts = [
            AggregatedAccount("111", "Bank1", "IFSC1", "Addr1", "", "", 5, "ACK1", 1000.0, 0.0, 50.0),
            AggregatedAccount("222", "Bank2", "IFSC2", "Addr2", "", "", 10, "ACK2", 2000.0, 0.0, 50.0),
            AggregatedAccount("333", "Bank3", "IFSC3", "Addr3", "", "", 3, "ACK3", 3000.0, 0.0, 50.0),
        ]
        
        results = dashboard.filter_by_min_transactions(accounts, 5)
        
        assert len(results) == 2
        assert all(acc.total_transactions >= 5 for acc in results)


class TestFilterByMinAmount:
    """Unit tests for filter_by_min_amount method."""
    
    def test_filters_correctly(self, dashboard):
        """Test filtering by minimum amount."""
        accounts = [
            AggregatedAccount("111", "Bank1", "IFSC1", "Addr1", "", "", 1, "ACK1", 1000.0, 0.0, 50.0),
            AggregatedAccount("222", "Bank2", "IFSC2", "Addr2", "", "", 1, "ACK2", 5000.0, 0.0, 50.0),
            AggregatedAccount("333", "Bank3", "IFSC3", "Addr3", "", "", 1, "ACK3", 3000.0, 0.0, 50.0),
        ]
        
        results = dashboard.filter_by_min_amount(accounts, 3000.0)
        
        assert len(results) == 2
        assert all(acc.total_amount >= 3000.0 for acc in results)


class TestGetFlaggedRows:
    """Unit tests for get_flagged_rows method."""
    
    def test_returns_flagged_accounts(self, dashboard):
        """Test that flagged accounts are returned correctly."""
        accounts = [
            AggregatedAccount("111111111", "Bank1", "IFSC1", "Addr1", "", "", 1, "ACK1", 1000.0, 0.0, 50.0),
            AggregatedAccount("222222222", "Bank2", "IFSC2", "Addr2", "", "", 1, "ACK2", 2000.0, 0.0, 50.0),
            AggregatedAccount("333333333", "Bank3", "IFSC3", "Addr3", "", "", 1, "ACK3", 3000.0, 0.0, 50.0),
        ]
        
        flagged = ["111111111", "333333333"]
        results = dashboard.get_flagged_rows(accounts, flagged)
        
        assert len(results) == 2
        assert results[0].account_number == "111111111"
        assert results[1].account_number == "333333333"
    
    def test_empty_flagged_list(self, dashboard):
        """Test with empty flagged list."""
        accounts = [
            AggregatedAccount("111111111", "Bank1", "IFSC1", "Addr1", "", "", 1, "ACK1", 1000.0, 0.0, 50.0),
        ]
        
        results = dashboard.get_flagged_rows(accounts, [])
        
        assert len(results) == 0
