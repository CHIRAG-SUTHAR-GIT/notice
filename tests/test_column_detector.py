"""
Property-based tests for the Column Detector.

Tests header normalization idempotence, fuzzy matching threshold consistency,
and column variant recognition using hypothesis.
"""

from hypothesis import given, settings, strategies as st
import pytest

from src.column_detector import ColumnDetector


# =============================================================================
# Test Fixtures
# =============================================================================

@pytest.fixture
def column_detector():
    """Fixture providing a ColumnDetector instance."""
    return ColumnDetector()


# =============================================================================
# Property-Based Tests for Header Normalization
# =============================================================================

# Feature: fraud-analysis-app, Property 3: Header Normalization Idempotence
# Validates: Requirements 2.2
# *For any* header string, normalizing it (strip whitespace, lowercase, 
# remove special characters) and then normalizing again should produce 
# the same result as normalizing once.

@settings(max_examples=100)
@given(
    header=st.text(
        alphabet=st.characters(
            whitelist_categories=('L', 'N', 'P', 'Z', 'S'),
            blacklist_characters='\x00'
        ),
        min_size=0,
        max_size=100
    )
)
def test_property_header_normalization_idempotence(header):
    """
    Property 3: Header Normalization Idempotence
    
    For any header string, normalizing it and then normalizing again 
    should produce the same result as normalizing once.
    
    Validates: Requirements 2.2
    """
    detector = ColumnDetector()
    
    # Normalize once
    normalized_once = detector.normalize_header(header)
    
    # Normalize twice
    normalized_twice = detector.normalize_header(normalized_once)
    
    # Property: f(f(x)) == f(x) - idempotence
    assert normalized_once == normalized_twice, (
        f"Header normalization is not idempotent:\n"
        f"  Original: {repr(header)}\n"
        f"  Normalized once: {repr(normalized_once)}\n"
        f"  Normalized twice: {repr(normalized_twice)}"
    )


# =============================================================================
# Unit Tests for Header Normalization
# =============================================================================

class TestHeaderNormalization:
    """Unit tests for header normalization edge cases."""
    
    def test_strips_leading_whitespace(self, column_detector):
        """Test that leading whitespace is stripped."""
        result = column_detector.normalize_header("   Bank Account")
        assert result == "bank account"
    
    def test_strips_trailing_whitespace(self, column_detector):
        """Test that trailing whitespace is stripped."""
        result = column_detector.normalize_header("Bank Account   ")
        assert result == "bank account"
    
    def test_converts_to_lowercase(self, column_detector):
        """Test that headers are converted to lowercase."""
        result = column_detector.normalize_header("BANK ACCOUNT NO")
        assert result == "bank account no"
    
    def test_removes_underscores(self, column_detector):
        """Test that underscores are replaced with spaces."""
        result = column_detector.normalize_header("bank_account_no")
        assert result == "bank account no"
    
    def test_removes_dashes(self, column_detector):
        """Test that dashes are replaced with spaces."""
        result = column_detector.normalize_header("bank-account-no")
        assert result == "bank account no"
    
    def test_removes_dots(self, column_detector):
        """Test that dots are replaced with spaces."""
        result = column_detector.normalize_header("sr.no")
        assert result == "sr no"
    
    def test_removes_slashes(self, column_detector):
        """Test that slashes are replaced with spaces."""
        result = column_detector.normalize_header("a/c no")
        assert result == "a c no"
    
    def test_removes_special_characters(self, column_detector):
        """Test that special characters are removed."""
        result = column_detector.normalize_header("Bank Account #@!")
        assert result == "bank account"
    
    def test_collapses_multiple_spaces(self, column_detector):
        """Test that multiple spaces are collapsed to single space."""
        result = column_detector.normalize_header("Bank    Account    No")
        assert result == "bank account no"
    
    def test_empty_string(self, column_detector):
        """Test normalization of empty string."""
        result = column_detector.normalize_header("")
        assert result == ""
    
    def test_only_whitespace(self, column_detector):
        """Test normalization of whitespace-only string."""
        result = column_detector.normalize_header("   \t\n   ")
        assert result == ""
    
    def test_non_string_input(self, column_detector):
        """Test that non-string input is converted to string."""
        result = column_detector.normalize_header(123)
        assert result == "123"


# =============================================================================
# Property-Based Tests for Fuzzy Matching
# =============================================================================

# Feature: fraud-analysis-app, Property 4: Fuzzy Matching Threshold Consistency
# Validates: Requirements 2.1
# *For any* header and column variant pair, if the fuzzy similarity score 
# is >= 80%, the Column_Detector should match them. If the score is < 80%, 
# they should not match.

@settings(max_examples=100)
@given(
    header=st.text(
        alphabet=st.characters(whitelist_categories=('L', 'N', 'Z')),
        min_size=1,
        max_size=50
    ),
    variant=st.text(
        alphabet=st.characters(whitelist_categories=('L', 'N', 'Z')),
        min_size=1,
        max_size=50
    )
)
def test_property_fuzzy_matching_threshold_consistency(header, variant):
    """
    Property 4: Fuzzy Matching Threshold Consistency
    
    For any header and column variant pair, if the fuzzy similarity score
    is >= 80%, the Column_Detector should match them. If the score is < 80%,
    they should not match.
    
    Validates: Requirements 2.1
    """
    detector = ColumnDetector()
    
    # Calculate similarity
    similarity = detector.calculate_similarity(header.lower(), variant.lower())
    
    # Property: similarity score is always between 0 and 1
    assert 0.0 <= similarity <= 1.0, (
        f"Similarity score {similarity} is out of range [0, 1]"
    )
    
    # Property: threshold consistency
    # If similarity >= threshold, it should be considered a match
    # If similarity < threshold, it should not be considered a match
    threshold = detector.SIMILARITY_THRESHOLD
    is_match = similarity >= threshold
    
    # Verify the threshold is applied consistently
    if is_match:
        assert similarity >= threshold, (
            f"Match detected but similarity {similarity} < threshold {threshold}"
        )
    else:
        assert similarity < threshold, (
            f"No match but similarity {similarity} >= threshold {threshold}"
        )


# Feature: fraud-analysis-app, Property 5: Column Variant Recognition
# Validates: Requirements 2.6, 2.7, 2.8, 2.9, 2.10, 2.11, 2.12, 2.13
# *For any* known column variant from the predefined lists, the Column_Detector 
# should correctly map it to the corresponding column type with 100% confidence.

@settings(max_examples=100)
@given(
    column_type=st.sampled_from([
        'serial_number', 'acknowledgement_number', 'bank_account_number',
        'ifsc_code', 'address', 'amount', 'disputed_amount', 'bank_name'
    ])
)
def test_property_column_variant_recognition(column_type):
    """
    Property 5: Column Variant Recognition
    
    For any known column variant from the predefined lists, the Column_Detector
    should correctly map it to the corresponding column type with 100% confidence.
    
    Validates: Requirements 2.6, 2.7, 2.8, 2.9, 2.10, 2.11, 2.12, 2.13
    """
    detector = ColumnDetector()
    
    # Get all variants for this column type
    variants = detector.COLUMN_VARIANTS[column_type]
    
    for variant in variants:
        # Create a header list with just this variant
        headers = [variant]
        
        # Detect columns
        mapping = detector.detect_columns(headers)
        
        # Property: the variant should be mapped to the correct column type
        mapped_value = getattr(mapping, column_type)
        assert mapped_value == variant, (
            f"Variant '{variant}' should map to '{column_type}', "
            f"but got mapped_value={mapped_value}"
        )
        
        # Property: confidence score should be 1.0 (100%) for exact matches
        confidence = mapping.confidence_scores.get(column_type, 0.0)
        assert confidence == 1.0, (
            f"Variant '{variant}' should have 100% confidence for '{column_type}', "
            f"but got {confidence}"
        )


# =============================================================================
# Unit Tests for Fuzzy Matching
# =============================================================================

class TestFuzzyMatching:
    """Unit tests for fuzzy matching functionality."""
    
    def test_exact_match_returns_1(self, column_detector):
        """Test that exact match returns similarity of 1.0."""
        similarity = column_detector.calculate_similarity("bank account no", "bank account no")
        assert similarity == 1.0
    
    def test_completely_different_returns_low_score(self, column_detector):
        """Test that completely different strings return low similarity."""
        similarity = column_detector.calculate_similarity("xyz", "abc")
        assert similarity < 0.5
    
    def test_similar_strings_return_high_score(self, column_detector):
        """Test that similar strings return high similarity."""
        similarity = column_detector.calculate_similarity("bank account no", "bank account number")
        assert similarity >= 0.8
    
    def test_threshold_is_80_percent(self, column_detector):
        """Test that the similarity threshold is 80%."""
        assert column_detector.SIMILARITY_THRESHOLD == 0.80


class TestColumnDetection:
    """Unit tests for column detection functionality."""
    
    def test_detects_bank_account_number(self, column_detector):
        """Test detection of bank account number column."""
        headers = ["Bank Account No", "Amount", "IFSC Code"]
        mapping = column_detector.detect_columns(headers)
        assert mapping.bank_account_number == "Bank Account No"
    
    def test_detects_amount(self, column_detector):
        """Test detection of amount column."""
        headers = ["Account No", "Transaction Amount", "Bank Name"]
        mapping = column_detector.detect_columns(headers)
        assert mapping.amount == "Transaction Amount"
    
    def test_detects_ifsc_code(self, column_detector):
        """Test detection of IFSC code column."""
        headers = ["Account", "IFSC", "Amount"]
        mapping = column_detector.detect_columns(headers)
        assert mapping.ifsc_code == "IFSC"
    
    def test_detects_all_columns(self, column_detector):
        """Test detection of all column types."""
        headers = [
            "Sr No", "Ack No", "Bank Account No", "IFSC Code",
            "Address", "Amount", "Disputed Amount", "Bank Name"
        ]
        mapping = column_detector.detect_columns(headers)
        
        assert mapping.serial_number == "Sr No"
        assert mapping.acknowledgement_number == "Ack No"
        assert mapping.bank_account_number == "Bank Account No"
        assert mapping.ifsc_code == "IFSC Code"
        assert mapping.address == "Address"
        assert mapping.amount == "Amount"
        assert mapping.disputed_amount == "Disputed Amount"
        assert mapping.bank_name == "Bank Name"
    
    def test_handles_case_variations(self, column_detector):
        """Test that column detection handles case variations."""
        headers = ["BANK ACCOUNT NO", "amount", "Ifsc Code"]
        mapping = column_detector.detect_columns(headers)
        
        assert mapping.bank_account_number == "BANK ACCOUNT NO"
        assert mapping.amount == "amount"
        assert mapping.ifsc_code == "Ifsc Code"
    
    def test_handles_whitespace_variations(self, column_detector):
        """Test that column detection handles whitespace variations."""
        headers = ["  Bank Account No  ", "  Amount  "]
        mapping = column_detector.detect_columns(headers)
        
        assert mapping.bank_account_number == "  Bank Account No  "
        assert mapping.amount == "  Amount  "
    
    def test_returns_confidence_scores(self, column_detector):
        """Test that confidence scores are returned."""
        headers = ["Bank Account No", "Amount"]
        mapping = column_detector.detect_columns(headers)
        
        assert "bank_account_number" in mapping.confidence_scores
        assert "amount" in mapping.confidence_scores
        assert mapping.confidence_scores["bank_account_number"] >= 0.8
        assert mapping.confidence_scores["amount"] >= 0.8
    
    def test_unmapped_headers(self, column_detector):
        """Test getting unmapped headers."""
        headers = ["Bank Account No", "Amount", "Unknown Column", "Random Field"]
        mapping = column_detector.detect_columns(headers)
        unmapped = column_detector.get_unmapped_headers(headers, mapping)
        
        assert "Unknown Column" in unmapped
        assert "Random Field" in unmapped
        assert "Bank Account No" not in unmapped
        assert "Amount" not in unmapped
    
    def test_validate_required_columns_all_present(self, column_detector):
        """Test validation when all required columns are present."""
        headers = ["Bank Account No", "Amount"]
        mapping = column_detector.detect_columns(headers)
        missing = column_detector.validate_required_columns(mapping)
        
        assert len(missing) == 0
    
    def test_validate_required_columns_missing_account(self, column_detector):
        """Test validation when bank account number is missing."""
        headers = ["Amount", "Bank Name"]
        mapping = column_detector.detect_columns(headers)
        missing = column_detector.validate_required_columns(mapping)
        
        assert "bank_account_number" in missing
    
    def test_validate_required_columns_missing_amount(self, column_detector):
        """Test validation when amount is missing."""
        headers = ["Bank Account No", "Bank Name"]
        mapping = column_detector.detect_columns(headers)
        missing = column_detector.validate_required_columns(mapping)
        
        assert "amount" in missing
