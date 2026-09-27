"""
Property-based tests for the Data Processor.

Tests empty row removal, whitespace trimming, account number standardization,
and amount parsing correctness using hypothesis.
"""

import pandas as pd
import pytest
from hypothesis import given, settings, strategies as st, assume

from src.data_processor import DataProcessor
from src.models import ColumnMapping


# =============================================================================
# Test Fixtures
# =============================================================================

@pytest.fixture
def data_processor():
    """Fixture providing a DataProcessor instance."""
    return DataProcessor()


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
# Property-Based Tests for Empty Row Removal
# =============================================================================

# Feature: fraud-analysis-app, Property 6: Empty Row Removal
# Validates: Requirements 3.1
# *For any* DataFrame, after processing, no row should exist where all cells 
# are empty or null.

@settings(max_examples=100)
@given(
    num_rows=st.integers(min_value=1, max_value=20),
    num_empty_rows=st.integers(min_value=0, max_value=10)
)
def test_property_empty_row_removal(num_rows, num_empty_rows):
    """
    Property 6: Empty Row Removal
    
    For any DataFrame, after processing, no row should exist where all cells
    are empty or null.
    
    Validates: Requirements 3.1
    """
    processor = DataProcessor()
    
    # Create a DataFrame with some valid rows and some empty rows
    data = []
    
    # Add valid rows
    for i in range(num_rows):
        data.append({
            'Account': f'12345678901{i}',
            'Amount': 1000.0 + i,
            'Name': f'Test{i}'
        })
    
    # Add empty rows (all None/empty)
    for _ in range(num_empty_rows):
        data.append({
            'Account': None,
            'Amount': None,
            'Name': None
        })
    
    # Also add rows with only whitespace
    for _ in range(num_empty_rows):
        data.append({
            'Account': '   ',
            'Amount': '  ',
            'Name': '\t\n'
        })
    
    df = pd.DataFrame(data)
    
    # Process the DataFrame
    result = processor.remove_empty_rows(df)
    
    # Property: No row should have all empty/null values
    for idx, row in result.iterrows():
        # Check if all values in the row are empty/null
        all_empty = all(
            pd.isna(val) or (isinstance(val, str) and val.strip() == '')
            for val in row
        )
        assert not all_empty, (
            f"Row {idx} has all empty values after processing: {row.to_dict()}"
        )
    
    # Property: All valid rows should be preserved
    assert len(result) == num_rows, (
        f"Expected {num_rows} valid rows, but got {len(result)}"
    )


# =============================================================================
# Property-Based Tests for Whitespace Trimming
# =============================================================================

# Feature: fraud-analysis-app, Property 7: Whitespace Trimming
# Validates: Requirements 3.2
# *For any* string cell in the DataFrame, after processing, it should have 
# no leading or trailing whitespace.

@settings(max_examples=100)
@given(
    content=st.text(
        alphabet=st.characters(whitelist_categories=('L', 'N')),
        min_size=1,
        max_size=20
    ),
    leading_ws=st.text(alphabet=' \t', min_size=0, max_size=5),
    trailing_ws=st.text(alphabet=' \t', min_size=0, max_size=5)
)
def test_property_whitespace_trimming(content, leading_ws, trailing_ws):
    """
    Property 7: Whitespace Trimming
    
    For any string cell in the DataFrame, after processing, it should have
    no leading or trailing whitespace.
    
    Validates: Requirements 3.2
    """
    processor = DataProcessor()
    
    # Create a string with leading and trailing whitespace
    padded_string = leading_ws + content + trailing_ws
    
    # Create a DataFrame with the padded string
    df = pd.DataFrame({
        'Column1': [padded_string],
        'Column2': [leading_ws + 'test' + trailing_ws]
    })
    
    # Process the DataFrame
    result = processor.trim_whitespace(df)
    
    # Property: No string cell should have leading or trailing whitespace
    for col in result.columns:
        for val in result[col]:
            if isinstance(val, str):
                assert val == val.strip(), (
                    f"Cell value '{repr(val)}' has leading or trailing whitespace"
                )


# =============================================================================
# Property-Based Tests for Account Number Standardization
# =============================================================================

# Feature: fraud-analysis-app, Property 8: Account Number Standardization
# Validates: Requirements 3.3
# *For any* bank account number string, after standardization, it should 
# contain no spaces or dashes.

@settings(max_examples=100)
@given(
    digits=st.text(alphabet='0123456789', min_size=9, max_size=18),
    separator=st.sampled_from([' ', '-', '  ', '--', ' - ', '']),
    chunk_size=st.integers(min_value=2, max_value=6)
)
def test_property_account_number_standardization(digits, separator, chunk_size):
    """
    Property 8: Account Number Standardization
    
    For any bank account number string, after standardization, it should
    contain no spaces or dashes.
    
    Validates: Requirements 3.3
    """
    processor = DataProcessor()
    
    # Create an account number with separators
    chunks = [digits[i:i+chunk_size] for i in range(0, len(digits), chunk_size)]
    formatted_account = separator.join(chunks)
    
    # Standardize the account number
    result = processor.standardize_account_number(formatted_account)
    
    # Property: Result should contain no spaces or dashes
    assert ' ' not in result, (
        f"Standardized account '{result}' contains spaces"
    )
    assert '-' not in result, (
        f"Standardized account '{result}' contains dashes"
    )
    
    # Property: All original digits should be preserved
    assert result == digits, (
        f"Standardized account '{result}' does not match original digits '{digits}'"
    )


# =============================================================================
# Property-Based Tests for Amount Parsing
# =============================================================================

# Feature: fraud-analysis-app, Property 11: Amount Parsing Correctness
# Validates: Requirements 3.6
# *For any* amount string containing currency symbols (₹, $, etc.) and/or commas,
# parsing should produce the correct numeric value equal to the string with 
# symbols and commas removed.

@settings(max_examples=100)
@given(
    amount=st.floats(min_value=0.01, max_value=1e8, allow_nan=False, allow_infinity=False),
    currency_symbol=st.sampled_from(['', '₹', '$', 'Rs.', 'Rs ', '£', '€']),
    use_commas=st.booleans()
)
def test_property_amount_parsing_correctness(amount, currency_symbol, use_commas):
    """
    Property 11: Amount Parsing Correctness
    
    For any amount string containing currency symbols and/or commas,
    parsing should produce the correct numeric value equal to the string
    with symbols and commas removed.
    
    Validates: Requirements 3.6
    """
    processor = DataProcessor()
    
    # Format the amount with optional commas and currency symbol
    if use_commas:
        amount_str = f"{currency_symbol}{amount:,.2f}"
    else:
        amount_str = f"{currency_symbol}{amount:.2f}"
    
    # Parse the amount
    result = processor.parse_amount(amount_str)
    
    # Property: Parsed value should be approximately equal to original
    # (using approximate comparison due to floating point precision)
    assert abs(result - amount) < 0.01, (
        f"Parsed amount {result} does not match original {amount} "
        f"from string '{amount_str}'"
    )


# =============================================================================
# Unit Tests for Empty Row Removal
# =============================================================================

class TestEmptyRowRemoval:
    """Unit tests for empty row removal functionality."""
    
    def test_removes_all_none_row(self, data_processor):
        """Test that rows with all None values are removed."""
        df = pd.DataFrame({
            'A': ['value1', None, 'value3'],
            'B': ['value2', None, 'value4']
        })
        result = data_processor.remove_empty_rows(df)
        assert len(result) == 2
    
    def test_removes_all_empty_string_row(self, data_processor):
        """Test that rows with all empty strings are removed."""
        df = pd.DataFrame({
            'A': ['value1', '', 'value3'],
            'B': ['value2', '', 'value4']
        })
        result = data_processor.remove_empty_rows(df)
        assert len(result) == 2
    
    def test_removes_all_whitespace_row(self, data_processor):
        """Test that rows with all whitespace are removed."""
        df = pd.DataFrame({
            'A': ['value1', '   ', 'value3'],
            'B': ['value2', '\t\n', 'value4']
        })
        result = data_processor.remove_empty_rows(df)
        assert len(result) == 2
    
    def test_keeps_partial_row(self, data_processor):
        """Test that rows with some values are kept."""
        df = pd.DataFrame({
            'A': ['value1', None, 'value3'],
            'B': ['value2', 'has_value', 'value4']
        })
        result = data_processor.remove_empty_rows(df)
        assert len(result) == 3
    
    def test_empty_dataframe(self, data_processor):
        """Test handling of empty DataFrame."""
        df = pd.DataFrame({'A': [], 'B': []})
        result = data_processor.remove_empty_rows(df)
        assert len(result) == 0


# =============================================================================
# Unit Tests for Whitespace Trimming
# =============================================================================

class TestWhitespaceTrimming:
    """Unit tests for whitespace trimming functionality."""
    
    def test_trims_leading_whitespace(self, data_processor):
        """Test that leading whitespace is trimmed."""
        df = pd.DataFrame({'A': ['   value']})
        result = data_processor.trim_whitespace(df)
        assert result['A'].iloc[0] == 'value'
    
    def test_trims_trailing_whitespace(self, data_processor):
        """Test that trailing whitespace is trimmed."""
        df = pd.DataFrame({'A': ['value   ']})
        result = data_processor.trim_whitespace(df)
        assert result['A'].iloc[0] == 'value'
    
    def test_trims_both_ends(self, data_processor):
        """Test that whitespace is trimmed from both ends."""
        df = pd.DataFrame({'A': ['   value   ']})
        result = data_processor.trim_whitespace(df)
        assert result['A'].iloc[0] == 'value'
    
    def test_preserves_internal_whitespace(self, data_processor):
        """Test that internal whitespace is preserved."""
        df = pd.DataFrame({'A': ['hello world']})
        result = data_processor.trim_whitespace(df)
        assert result['A'].iloc[0] == 'hello world'
    
    def test_handles_non_string_columns(self, data_processor):
        """Test that non-string columns are not affected."""
        df = pd.DataFrame({'A': [1, 2, 3], 'B': ['  text  ', '  more  ', '  data  ']})
        result = data_processor.trim_whitespace(df)
        assert list(result['A']) == [1, 2, 3]
        assert result['B'].iloc[0] == 'text'


# =============================================================================
# Unit Tests for Account Number Standardization
# =============================================================================

class TestAccountNumberStandardization:
    """Unit tests for account number standardization."""
    
    def test_removes_spaces(self, data_processor):
        """Test that spaces are removed."""
        result = data_processor.standardize_account_number('1234 5678 9012')
        assert result == '123456789012'
    
    def test_removes_dashes(self, data_processor):
        """Test that dashes are removed."""
        result = data_processor.standardize_account_number('1234-5678-9012')
        assert result == '123456789012'
    
    def test_removes_mixed_separators(self, data_processor):
        """Test that mixed separators are removed."""
        result = data_processor.standardize_account_number('1234 5678-9012')
        assert result == '123456789012'
    
    def test_handles_none(self, data_processor):
        """Test handling of None input."""
        result = data_processor.standardize_account_number(None)
        assert result == ''
    
    def test_handles_nan_string(self, data_processor):
        """Test handling of 'nan' string."""
        result = data_processor.standardize_account_number('nan')
        assert result == ''


# =============================================================================
# Unit Tests for Amount Parsing
# =============================================================================

class TestAmountParsing:
    """Unit tests for amount parsing functionality."""
    
    def test_parses_plain_number(self, data_processor):
        """Test parsing of plain number."""
        result = data_processor.parse_amount('1000.00')
        assert result == 1000.00
    
    def test_parses_with_rupee_symbol(self, data_processor):
        """Test parsing with rupee symbol."""
        result = data_processor.parse_amount('₹1000.00')
        assert result == 1000.00
    
    def test_parses_with_dollar_symbol(self, data_processor):
        """Test parsing with dollar symbol."""
        result = data_processor.parse_amount('$1000.00')
        assert result == 1000.00
    
    def test_parses_with_commas(self, data_processor):
        """Test parsing with comma separators."""
        result = data_processor.parse_amount('1,000,000.00')
        assert result == 1000000.00
    
    def test_parses_with_symbol_and_commas(self, data_processor):
        """Test parsing with both symbol and commas."""
        result = data_processor.parse_amount('₹1,00,000.00')
        assert result == 100000.00
    
    def test_parses_rs_prefix(self, data_processor):
        """Test parsing with Rs. prefix."""
        result = data_processor.parse_amount('Rs. 5000')
        assert result == 5000.00
    
    def test_handles_none(self, data_processor):
        """Test handling of None input."""
        result = data_processor.parse_amount(None)
        assert result == 0.0
    
    def test_handles_empty_string(self, data_processor):
        """Test handling of empty string."""
        result = data_processor.parse_amount('')
        assert result == 0.0
    
    def test_handles_invalid_format(self, data_processor):
        """Test handling of invalid format."""
        result = data_processor.parse_amount('not a number')
        assert result == 0.0


# =============================================================================
# Unit Tests for Clean DataFrame
# =============================================================================

class TestCleanDataFrame:
    """Unit tests for the clean_dataframe orchestration method."""
    
    def test_applies_all_cleaning_operations(self, data_processor, sample_column_mapping):
        """Test that all cleaning operations are applied."""
        df = pd.DataFrame({
            'Sr No': [1, None, 3],
            'Ack No': ['  ACK001  ', None, '  ACK003  '],
            'Bank Account No': ['1234 5678 9012', None, '9876-5432-1098'],
            'IFSC Code': ['SBIN0001234', None, 'HDFC0005678'],
            'Address': ['  Address 1  ', None, '  Address 3  '],
            'Amount': ['₹10,000.00', None, '$25,000.00'],
            'Disputed Amount': ['Rs. 5000', None, '15000'],
            'Bank Name': ['  SBI  ', None, '  HDFC  ']
        })
        
        result = data_processor.clean_dataframe(df, sample_column_mapping)
        
        # Check empty rows removed
        assert len(result) == 2
        
        # Check whitespace trimmed
        assert result['Ack No'].iloc[0] == 'ACK001'
        assert result['Bank Name'].iloc[0] == 'SBI'
        
        # Check account numbers standardized
        assert result['Bank Account No'].iloc[0] == '123456789012'
        assert result['Bank Account No'].iloc[1] == '987654321098'
        
        # Check amounts parsed
        assert result['Amount'].iloc[0] == 10000.00
        assert result['Amount'].iloc[1] == 25000.00
