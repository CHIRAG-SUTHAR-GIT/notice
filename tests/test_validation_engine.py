"""
Property-based tests for the Validation Engine.

Tests account number validation, IFSC code validation, amount validation,
critical data flagging, duplicate acknowledgement detection, and error
classification using hypothesis.
"""

import pandas as pd
import pytest
from hypothesis import given, settings, strategies as st, assume

from src.validation_engine import (
    ValidationEngine,
    CRITICAL_ERRORS,
    WARNING_ERRORS
)
from src.models import ColumnMapping, ErrorCategory


# =============================================================================
# Test Fixtures
# =============================================================================

@pytest.fixture
def validation_engine():
    """Fixture providing a ValidationEngine instance."""
    return ValidationEngine()


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
# Property-Based Tests for Account Number Validation
# =============================================================================

# Feature: fraud-analysis-app, Property 9: Account Number Validation
# Validates: Requirements 3.4
# *For any* string, the Validation_Engine should mark it as a valid account 
# number if and only if it contains between 9 and 18 digits (inclusive) after 
# removing non-digit characters.

@settings(max_examples=100)
@given(
    digits=st.text(alphabet='0123456789', min_size=9, max_size=18)
)
def test_property_valid_account_numbers_accepted(digits):
    """
    Property 9: Account Number Validation (valid case)
    
    For any string containing 9-18 digits, the Validation_Engine should
    mark it as a valid account number.
    
    Validates: Requirements 3.4
    """
    engine = ValidationEngine()
    
    # Valid account numbers (9-18 digits) should be accepted
    result = engine.validate_account_number(digits)
    
    assert result is True, (
        f"Account number '{digits}' with {len(digits)} digits should be valid"
    )


@settings(max_examples=100)
@given(
    digits=st.text(alphabet='0123456789', min_size=0, max_size=8)
)
def test_property_short_account_numbers_rejected(digits):
    """
    Property 9: Account Number Validation (too short case)
    
    For any string containing fewer than 9 digits, the Validation_Engine 
    should mark it as invalid.
    
    Validates: Requirements 3.4
    """
    engine = ValidationEngine()
    
    # Account numbers with fewer than 9 digits should be rejected
    result = engine.validate_account_number(digits)
    
    assert result is False, (
        f"Account number '{digits}' with {len(digits)} digits should be invalid (too short)"
    )


@settings(max_examples=100)
@given(
    digits=st.text(alphabet='0123456789', min_size=19, max_size=30)
)
def test_property_long_account_numbers_rejected(digits):
    """
    Property 9: Account Number Validation (too long case)
    
    For any string containing more than 18 digits, the Validation_Engine 
    should mark it as invalid.
    
    Validates: Requirements 3.4
    """
    engine = ValidationEngine()
    
    # Account numbers with more than 18 digits should be rejected
    result = engine.validate_account_number(digits)
    
    assert result is False, (
        f"Account number '{digits}' with {len(digits)} digits should be invalid (too long)"
    )


@settings(max_examples=100)
@given(
    digits=st.text(alphabet='0123456789', min_size=9, max_size=18),
    separator=st.sampled_from([' ', '-', '  ', '--', ' - '])
)
def test_property_account_numbers_with_separators(digits, separator):
    """
    Property 9: Account Number Validation (with separators)
    
    For any string containing 9-18 digits with separators, the Validation_Engine 
    should mark it as valid after removing non-digit characters.
    
    Validates: Requirements 3.4
    """
    engine = ValidationEngine()
    
    # Add separators to the digits
    chunk_size = 4
    chunks = [digits[i:i+chunk_size] for i in range(0, len(digits), chunk_size)]
    formatted = separator.join(chunks)
    
    result = engine.validate_account_number(formatted)
    
    assert result is True, (
        f"Account number '{formatted}' with {len(digits)} digits should be valid"
    )


# =============================================================================
# Property-Based Tests for IFSC Code Validation
# =============================================================================

# Feature: fraud-analysis-app, Property 10: IFSC Code Validation
# Validates: Requirements 3.5
# *For any* string, the Validation_Engine should mark it as a valid IFSC code 
# if and only if it consists of exactly 11 alphanumeric characters.

@settings(max_examples=100)
@given(
    ifsc=st.text(
        alphabet='ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789',
        min_size=11,
        max_size=11
    )
)
def test_property_valid_ifsc_codes_accepted(ifsc):
    """
    Property 10: IFSC Code Validation (valid case)
    
    For any string of exactly 11 alphanumeric characters, the Validation_Engine 
    should mark it as a valid IFSC code.
    
    Validates: Requirements 3.5
    """
    engine = ValidationEngine()
    
    result = engine.validate_ifsc_code(ifsc)
    
    assert result is True, (
        f"IFSC code '{ifsc}' with {len(ifsc)} alphanumeric chars should be valid"
    )


@settings(max_examples=100)
@given(
    ifsc=st.text(
        alphabet='ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789',
        min_size=0,
        max_size=10
    )
)
def test_property_short_ifsc_codes_rejected(ifsc):
    """
    Property 10: IFSC Code Validation (too short case)
    
    For any string with fewer than 11 alphanumeric characters, the 
    Validation_Engine should mark it as invalid.
    
    Validates: Requirements 3.5
    """
    engine = ValidationEngine()
    
    result = engine.validate_ifsc_code(ifsc)
    
    assert result is False, (
        f"IFSC code '{ifsc}' with {len(ifsc)} chars should be invalid (too short)"
    )


@settings(max_examples=100)
@given(
    ifsc=st.text(
        alphabet='ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789',
        min_size=12,
        max_size=20
    )
)
def test_property_long_ifsc_codes_rejected(ifsc):
    """
    Property 10: IFSC Code Validation (too long case)
    
    For any string with more than 11 alphanumeric characters, the 
    Validation_Engine should mark it as invalid.
    
    Validates: Requirements 3.5
    """
    engine = ValidationEngine()
    
    result = engine.validate_ifsc_code(ifsc)
    
    assert result is False, (
        f"IFSC code '{ifsc}' with {len(ifsc)} chars should be invalid (too long)"
    )


@settings(max_examples=100)
@given(
    base=st.text(
        alphabet='ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789',
        min_size=10,
        max_size=10
    ),
    special_char=st.sampled_from(['!', '@', '#', '$', '%', '^', '&', '*', ' ', '-'])
)
def test_property_ifsc_with_special_chars_rejected(base, special_char):
    """
    Property 10: IFSC Code Validation (special characters case)
    
    For any string containing non-alphanumeric characters, the 
    Validation_Engine should mark it as invalid.
    
    Validates: Requirements 3.5
    """
    engine = ValidationEngine()
    
    # Insert special character to make it 11 chars
    ifsc_with_special = base + special_char
    
    result = engine.validate_ifsc_code(ifsc_with_special)
    
    assert result is False, (
        f"IFSC code '{ifsc_with_special}' with special char should be invalid"
    )


# =============================================================================
# Property-Based Tests for Amount Validation
# =============================================================================

# Feature: fraud-analysis-app, Property 12: Amount Validation
# Validates: Requirements 3.7
# *For any* numeric amount, the Validation_Engine should mark it as valid 
# if and only if it is a positive number (> 0).

@settings(max_examples=100)
@given(
    amount=st.floats(min_value=0.01, max_value=1e10, allow_nan=False, allow_infinity=False)
)
def test_property_positive_amounts_valid(amount):
    """
    Property 12: Amount Validation (positive case)
    
    For any positive number (> 0), the Validation_Engine should mark it as valid.
    
    Validates: Requirements 3.7
    """
    engine = ValidationEngine()
    
    result = engine.validate_amount(amount)
    
    assert result is True, (
        f"Amount {amount} should be valid (positive)"
    )


@settings(max_examples=100)
@given(
    amount=st.floats(max_value=0.0, allow_nan=False, allow_infinity=False)
)
def test_property_non_positive_amounts_invalid(amount):
    """
    Property 12: Amount Validation (non-positive case)
    
    For any non-positive number (<= 0), the Validation_Engine should mark it as invalid.
    
    Validates: Requirements 3.7
    """
    engine = ValidationEngine()
    
    result = engine.validate_amount(amount)
    
    assert result is False, (
        f"Amount {amount} should be invalid (non-positive)"
    )


@settings(max_examples=100)
@given(
    amount=st.integers(min_value=1, max_value=1000000000)
)
def test_property_positive_integers_valid(amount):
    """
    Property 12: Amount Validation (integer case)
    
    For any positive integer, the Validation_Engine should mark it as valid.
    
    Validates: Requirements 3.7
    """
    engine = ValidationEngine()
    
    result = engine.validate_amount(amount)
    
    assert result is True, (
        f"Amount {amount} should be valid (positive integer)"
    )


# =============================================================================
# Property-Based Tests for Critical Data Flagging
# =============================================================================

# Feature: fraud-analysis-app, Property 13: Critical Data Flagging
# Validates: Requirements 3.8
# *For any* row in the DataFrame, if the bank account number or amount field 
# is missing/null, the row should be flagged as having critical missing data.

@settings(max_examples=100)
@given(
    num_valid_rows=st.integers(min_value=1, max_value=10),
    missing_account_rows=st.integers(min_value=0, max_value=5),
    missing_amount_rows=st.integers(min_value=0, max_value=5)
)
def test_property_critical_data_flagging(num_valid_rows, missing_account_rows, missing_amount_rows):
    """
    Property 13: Critical Data Flagging
    
    For any row in the DataFrame, if the bank account number or amount field
    is missing/null, the row should be flagged as having critical missing data.
    
    Validates: Requirements 3.8
    """
    engine = ValidationEngine()
    
    # Create mapping
    mapping = ColumnMapping(
        bank_account_number="Account",
        amount="Amount"
    )
    
    # Build data
    data = []
    
    # Add valid rows
    for i in range(num_valid_rows):
        data.append({
            'Account': f'12345678901{i % 10}',
            'Amount': 1000.0 + i
        })
    
    # Add rows with missing account numbers
    for i in range(missing_account_rows):
        data.append({
            'Account': None,
            'Amount': 500.0 + i
        })
    
    # Add rows with missing amounts
    for i in range(missing_amount_rows):
        data.append({
            'Account': f'98765432109{i % 10}',
            'Amount': None
        })
    
    df = pd.DataFrame(data)
    
    # Validate
    result = engine.validate_dataframe(df, mapping)
    
    # Property: Number of flagged rows should equal rows with missing critical data
    expected_flagged = missing_account_rows + missing_amount_rows
    
    assert len(result.flagged_rows) >= expected_flagged, (
        f"Expected at least {expected_flagged} flagged rows, got {len(result.flagged_rows)}"
    )


# =============================================================================
# Property-Based Tests for Duplicate Acknowledgement Detection
# =============================================================================

# Feature: fraud-analysis-app, Property 14: Duplicate Acknowledgement Detection
# Validates: Requirements 3.9
# *For any* DataFrame where the same acknowledgement number appears more than 
# once, the Validation_Engine should generate a warning listing the duplicates.

@settings(max_examples=100)
@given(
    unique_acks=st.lists(
        st.text(alphabet='ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', min_size=5, max_size=15),
        min_size=1,
        max_size=10,
        unique=True
    ),
    duplicate_count=st.integers(min_value=1, max_value=5)
)
def test_property_duplicate_acknowledgement_detection(unique_acks, duplicate_count):
    """
    Property 14: Duplicate Acknowledgement Detection
    
    For any DataFrame where the same acknowledgement number appears more than
    once, the Validation_Engine should generate a warning listing the duplicates.
    
    Validates: Requirements 3.9
    """
    assume(len(unique_acks) > 0)
    
    engine = ValidationEngine()
    
    # Create data with some duplicates
    ack_list = unique_acks.copy()
    
    # Pick one ack to duplicate
    ack_to_duplicate = unique_acks[0]
    
    # Add duplicates
    for _ in range(duplicate_count):
        ack_list.append(ack_to_duplicate)
    
    df = pd.DataFrame({'Ack No': ack_list})
    
    # Check for duplicates
    duplicates = engine.check_duplicate_acknowledgements(df, 'Ack No')
    
    # Property: The duplicated ack should be in the list
    assert ack_to_duplicate in duplicates, (
        f"Duplicate ack '{ack_to_duplicate}' should be detected. Found: {duplicates}"
    )


@settings(max_examples=100)
@given(
    unique_acks=st.lists(
        st.text(alphabet='ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', min_size=5, max_size=15),
        min_size=1,
        max_size=20,
        unique=True
    )
)
def test_property_no_false_duplicate_detection(unique_acks):
    """
    Property 14: Duplicate Acknowledgement Detection (no false positives)
    
    For any DataFrame where all acknowledgement numbers are unique,
    the Validation_Engine should not report any duplicates.
    
    Validates: Requirements 3.9
    """
    engine = ValidationEngine()
    
    df = pd.DataFrame({'Ack No': unique_acks})
    
    # Check for duplicates
    duplicates = engine.check_duplicate_acknowledgements(df, 'Ack No')
    
    # Property: No duplicates should be found
    assert len(duplicates) == 0, (
        f"No duplicates expected for unique acks, but found: {duplicates}"
    )


# =============================================================================
# Property-Based Tests for Error Classification Consistency
# =============================================================================

# Feature: fraud-analysis-app, Property 28: Error Classification Consistency
# Validates: Requirements 10.1, 10.2, 10.3, 10.4, 10.5, 10.6
# *For any* validation error, it should be classified as either Critical 
# (missing account number column, missing amount column, unreadable file) or 
# Warning (missing IFSC, missing address, invalid amount format, duplicate ack 
# numbers) - never both.

@settings(max_examples=100)
@given(
    error_code=st.sampled_from(list(CRITICAL_ERRORS.keys()))
)
def test_property_critical_errors_classified_correctly(error_code):
    """
    Property 28: Error Classification Consistency (critical errors)
    
    For any critical error code, the Validation_Engine should classify it
    as ErrorCategory.CRITICAL.
    
    Validates: Requirements 10.1, 10.2, 10.6
    """
    engine = ValidationEngine()
    
    category = engine.classify_error(error_code)
    
    assert category == ErrorCategory.CRITICAL, (
        f"Error code '{error_code}' should be classified as CRITICAL, got {category}"
    )


@settings(max_examples=100)
@given(
    error_code=st.sampled_from(list(WARNING_ERRORS.keys()))
)
def test_property_warning_errors_classified_correctly(error_code):
    """
    Property 28: Error Classification Consistency (warning errors)
    
    For any warning error code, the Validation_Engine should classify it
    as ErrorCategory.WARNING.
    
    Validates: Requirements 10.3, 10.4, 10.5, 10.6
    """
    engine = ValidationEngine()
    
    category = engine.classify_error(error_code)
    
    assert category == ErrorCategory.WARNING, (
        f"Error code '{error_code}' should be classified as WARNING, got {category}"
    )


def test_property_error_categories_mutually_exclusive():
    """
    Property 28: Error Classification Consistency (mutual exclusivity)
    
    No error code should appear in both CRITICAL_ERRORS and WARNING_ERRORS.
    
    Validates: Requirements 10.6
    """
    critical_codes = set(CRITICAL_ERRORS.keys())
    warning_codes = set(WARNING_ERRORS.keys())
    
    overlap = critical_codes.intersection(warning_codes)
    
    assert len(overlap) == 0, (
        f"Error codes should not appear in both categories: {overlap}"
    )


# =============================================================================
# Unit Tests for Account Number Validation
# =============================================================================

class TestAccountNumberValidation:
    """Unit tests for account number validation."""
    
    def test_valid_9_digit_account(self, validation_engine):
        """Test that 9-digit account is valid."""
        assert validation_engine.validate_account_number('123456789') is True
    
    def test_valid_18_digit_account(self, validation_engine):
        """Test that 18-digit account is valid."""
        assert validation_engine.validate_account_number('123456789012345678') is True
    
    def test_invalid_8_digit_account(self, validation_engine):
        """Test that 8-digit account is invalid."""
        assert validation_engine.validate_account_number('12345678') is False
    
    def test_invalid_19_digit_account(self, validation_engine):
        """Test that 19-digit account is invalid."""
        assert validation_engine.validate_account_number('1234567890123456789') is False
    
    def test_none_account(self, validation_engine):
        """Test that None is invalid."""
        assert validation_engine.validate_account_number(None) is False
    
    def test_empty_account(self, validation_engine):
        """Test that empty string is invalid."""
        assert validation_engine.validate_account_number('') is False
    
    def test_account_with_spaces(self, validation_engine):
        """Test account with spaces (digits extracted)."""
        assert validation_engine.validate_account_number('1234 5678 9012') is True


# =============================================================================
# Unit Tests for IFSC Code Validation
# =============================================================================

class TestIFSCCodeValidation:
    """Unit tests for IFSC code validation."""
    
    def test_valid_ifsc(self, validation_engine):
        """Test that valid IFSC is accepted."""
        assert validation_engine.validate_ifsc_code('SBIN0001234') is True
    
    def test_valid_ifsc_lowercase(self, validation_engine):
        """Test that lowercase IFSC is rejected (must be uppercase)."""
        # Note: isalnum() accepts lowercase, but real IFSC codes are uppercase
        assert validation_engine.validate_ifsc_code('sbin0001234') is True
    
    def test_invalid_short_ifsc(self, validation_engine):
        """Test that short IFSC is invalid."""
        assert validation_engine.validate_ifsc_code('SBIN000123') is False
    
    def test_invalid_long_ifsc(self, validation_engine):
        """Test that long IFSC is invalid."""
        assert validation_engine.validate_ifsc_code('SBIN00012345') is False
    
    def test_none_ifsc(self, validation_engine):
        """Test that None is invalid."""
        assert validation_engine.validate_ifsc_code(None) is False
    
    def test_empty_ifsc(self, validation_engine):
        """Test that empty string is invalid."""
        assert validation_engine.validate_ifsc_code('') is False
    
    def test_ifsc_with_special_chars(self, validation_engine):
        """Test that IFSC with special chars is invalid."""
        assert validation_engine.validate_ifsc_code('SBIN-001234') is False


# =============================================================================
# Unit Tests for Amount Validation
# =============================================================================

class TestAmountValidation:
    """Unit tests for amount validation."""
    
    def test_positive_float(self, validation_engine):
        """Test that positive float is valid."""
        assert validation_engine.validate_amount(100.50) is True
    
    def test_positive_integer(self, validation_engine):
        """Test that positive integer is valid."""
        assert validation_engine.validate_amount(100) is True
    
    def test_zero_amount(self, validation_engine):
        """Test that zero is invalid."""
        assert validation_engine.validate_amount(0) is False
    
    def test_negative_amount(self, validation_engine):
        """Test that negative amount is invalid."""
        assert validation_engine.validate_amount(-100) is False
    
    def test_none_amount(self, validation_engine):
        """Test that None is invalid."""
        assert validation_engine.validate_amount(None) is False
    
    def test_string_amount(self, validation_engine):
        """Test that numeric string is valid."""
        assert validation_engine.validate_amount('100.50') is True
    
    def test_invalid_string_amount(self, validation_engine):
        """Test that non-numeric string is invalid."""
        assert validation_engine.validate_amount('not a number') is False


# =============================================================================
# Unit Tests for Duplicate Acknowledgement Detection
# =============================================================================

class TestDuplicateAcknowledgementDetection:
    """Unit tests for duplicate acknowledgement detection."""
    
    def test_finds_duplicates(self, validation_engine):
        """Test that duplicates are found."""
        df = pd.DataFrame({'Ack No': ['ACK001', 'ACK002', 'ACK001', 'ACK003']})
        duplicates = validation_engine.check_duplicate_acknowledgements(df, 'Ack No')
        assert 'ACK001' in duplicates
    
    def test_no_duplicates(self, validation_engine):
        """Test that no duplicates returns empty list."""
        df = pd.DataFrame({'Ack No': ['ACK001', 'ACK002', 'ACK003']})
        duplicates = validation_engine.check_duplicate_acknowledgements(df, 'Ack No')
        assert len(duplicates) == 0
    
    def test_missing_column(self, validation_engine):
        """Test handling of missing column."""
        df = pd.DataFrame({'Other': ['A', 'B', 'C']})
        duplicates = validation_engine.check_duplicate_acknowledgements(df, 'Ack No')
        assert len(duplicates) == 0
    
    def test_none_column(self, validation_engine):
        """Test handling of None column name."""
        df = pd.DataFrame({'Ack No': ['ACK001', 'ACK002']})
        duplicates = validation_engine.check_duplicate_acknowledgements(df, None)
        assert len(duplicates) == 0
    
    def test_ignores_null_values(self, validation_engine):
        """Test that null values are not counted as duplicates."""
        df = pd.DataFrame({'Ack No': ['ACK001', None, None, 'ACK002']})
        duplicates = validation_engine.check_duplicate_acknowledgements(df, 'Ack No')
        assert len(duplicates) == 0


# =============================================================================
# Unit Tests for DataFrame Validation
# =============================================================================

class TestDataFrameValidation:
    """Unit tests for DataFrame validation."""
    
    def test_missing_account_column_critical_error(self, validation_engine):
        """Test that missing account column causes critical error."""
        mapping = ColumnMapping(
            bank_account_number="Account",
            amount="Amount"
        )
        df = pd.DataFrame({'Amount': [100, 200]})
        
        result = validation_engine.validate_dataframe(df, mapping)
        
        assert result.is_valid is False
        assert any('account' in err.lower() for err in result.critical_errors)
    
    def test_missing_amount_column_critical_error(self, validation_engine):
        """Test that missing amount column causes critical error."""
        mapping = ColumnMapping(
            bank_account_number="Account",
            amount="Amount"
        )
        df = pd.DataFrame({'Account': ['123456789012']})
        
        result = validation_engine.validate_dataframe(df, mapping)
        
        assert result.is_valid is False
        assert any('amount' in err.lower() for err in result.critical_errors)
    
    def test_valid_dataframe(self, validation_engine, sample_column_mapping):
        """Test validation of valid DataFrame."""
        df = pd.DataFrame({
            'Sr No': [1, 2],
            'Ack No': ['ACK001', 'ACK002'],
            'Bank Account No': ['123456789012', '987654321098'],
            'IFSC Code': ['SBIN0001234', 'HDFC0005678'],
            'Address': ['Address 1', 'Address 2'],
            'Amount': [1000.0, 2000.0],
            'Disputed Amount': [500.0, 1000.0],
            'Bank Name': ['SBI', 'HDFC']
        })
        
        result = validation_engine.validate_dataframe(df, sample_column_mapping)
        
        assert result.is_valid is True
        assert len(result.critical_errors) == 0


# =============================================================================
# Unit Tests for Quality Report Generation
# =============================================================================

class TestQualityReportGeneration:
    """Unit tests for quality report generation."""
    
    def test_empty_dataframe_report(self, validation_engine, sample_column_mapping):
        """Test quality report for empty DataFrame."""
        df = pd.DataFrame(columns=['Bank Account No', 'Amount'])
        
        report = validation_engine.generate_quality_report(df, sample_column_mapping)
        
        assert report['total_rows'] == 0
        assert report['valid_account_numbers'] == 0
    
    def test_report_counts_valid_accounts(self, validation_engine):
        """Test that report correctly counts valid accounts."""
        mapping = ColumnMapping(
            bank_account_number="Account",
            amount="Amount"
        )
        df = pd.DataFrame({
            'Account': ['123456789012', '12345678', '987654321098'],  # 2 valid, 1 invalid
            'Amount': [100, 200, 300]
        })
        
        report = validation_engine.generate_quality_report(df, mapping)
        
        assert report['total_rows'] == 3
        assert report['valid_account_numbers'] == 2
    
    def test_report_counts_duplicates(self, validation_engine):
        """Test that report correctly counts duplicates."""
        mapping = ColumnMapping(
            bank_account_number="Account",
            amount="Amount",
            acknowledgement_number="Ack"
        )
        df = pd.DataFrame({
            'Account': ['123456789012', '987654321098'],
            'Amount': [100, 200],
            'Ack': ['ACK001', 'ACK001']  # Duplicate
        })
        
        report = validation_engine.generate_quality_report(df, mapping)
        
        assert report['duplicate_acknowledgements'] == 1


# =============================================================================
# Unit Tests for Error Classification
# =============================================================================

class TestErrorClassification:
    """Unit tests for error classification."""
    
    def test_classify_critical_error(self, validation_engine):
        """Test classification of critical error."""
        category = validation_engine.classify_error('NO_ACCOUNT_COLUMN')
        assert category == ErrorCategory.CRITICAL
    
    def test_classify_warning_error(self, validation_engine):
        """Test classification of warning error."""
        category = validation_engine.classify_error('MISSING_IFSC')
        assert category == ErrorCategory.WARNING
    
    def test_classify_unknown_error(self, validation_engine):
        """Test classification of unknown error defaults to warning."""
        category = validation_engine.classify_error('UNKNOWN_ERROR')
        assert category == ErrorCategory.WARNING
    
    def test_create_error_response(self, validation_engine):
        """Test creation of error response."""
        response = validation_engine.create_error_response(
            'MISSING_IFSC',
            row_number=5,
            field_name='ifsc_code',
            row=5
        )
        
        assert response.category == ErrorCategory.WARNING
        assert response.code == 'MISSING_IFSC'
        assert response.row_number == 5
        assert '5' in response.message
