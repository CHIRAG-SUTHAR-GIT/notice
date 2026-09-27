"""
Property-based tests for the Upload Service.

Tests file type validation and file size validation using hypothesis.
"""

import io
from hypothesis import given, settings, strategies as st
import pytest

from src.upload_service import UploadService, FileValidationResult


# =============================================================================
# Test Fixtures
# =============================================================================

@pytest.fixture
def upload_service():
    """Fixture providing an UploadService instance."""
    return UploadService()


# =============================================================================
# Property-Based Tests
# =============================================================================

# Feature: fraud-analysis-app, Property 1: File Type Validation
# Validates: Requirements 1.1
# *For any* file upload attempt, the Upload_Service should accept the file 
# if and only if its extension is in the allowed set (.xlsx, .xls, .csv).

@settings(max_examples=100)
@given(
    extension=st.sampled_from(['.xlsx', '.xls', '.csv', '.txt', '.pdf', '.doc', '.json', '.xml', '.zip', '.exe'])
)
def test_property_file_type_validation(extension):
    """
    Property 1: File Type Validation
    
    For any file upload attempt, the Upload_Service should accept the file
    if and only if its extension is in the allowed set (.xlsx, .xls, .csv).
    
    Validates: Requirements 1.1
    """
    service = UploadService()
    
    # Create a small dummy file
    file_content = b"dummy content for testing"
    file = io.BytesIO(file_content)
    filename = f"test_file{extension}"
    
    result = service.validate_file(file, filename)
    
    # Property: file is valid iff extension is in allowed set
    expected_valid = extension in service.ALLOWED_EXTENSIONS
    assert result.is_valid == expected_valid, (
        f"Extension {extension}: expected is_valid={expected_valid}, got {result.is_valid}"
    )
    
    # If invalid, should have error message
    if not expected_valid:
        assert result.error_message is not None
        assert "Invalid file type" in result.error_message


# Feature: fraud-analysis-app, Property 2: File Size Validation
# Validates: Requirements 1.2
# *For any* file with size greater than 50MB, the Upload_Service should reject 
# the upload. *For any* file with size less than or equal to 50MB, the 
# Upload_Service should not reject based on size alone.

@settings(max_examples=100)
@given(
    file_size_mb=st.floats(min_value=0.001, max_value=100.0, allow_nan=False, allow_infinity=False)
)
def test_property_file_size_validation(file_size_mb):
    """
    Property 2: File Size Validation
    
    For any file with size greater than 50MB, the Upload_Service should reject
    the upload. For any file with size less than or equal to 50MB, the
    Upload_Service should not reject based on size alone.
    
    Validates: Requirements 1.2
    """
    service = UploadService()
    
    # Create a file of the specified size (approximately)
    # We use a valid extension to isolate size validation
    file_size_bytes = int(file_size_mb * 1024 * 1024)
    file_content = b"x" * file_size_bytes
    file = io.BytesIO(file_content)
    filename = "test_file.xlsx"
    
    result = service.validate_file(file, filename)
    
    # Property: file is rejected if size > 50MB
    if file_size_mb > service.MAX_FILE_SIZE_MB:
        assert not result.is_valid, (
            f"File of {file_size_mb}MB should be rejected (max: {service.MAX_FILE_SIZE_MB}MB)"
        )
        assert result.error_message is not None
        assert "exceeds maximum" in result.error_message.lower() or "size" in result.error_message.lower()
    else:
        # File should be accepted (size is within limit)
        assert result.is_valid, (
            f"File of {file_size_mb}MB should be accepted (max: {service.MAX_FILE_SIZE_MB}MB)"
        )


# =============================================================================
# Unit Tests for Edge Cases and Specific Examples
# =============================================================================

class TestUploadServiceValidation:
    """Unit tests for UploadService validation methods."""
    
    def test_valid_xlsx_file(self, upload_service):
        """Test that .xlsx files are accepted."""
        file = io.BytesIO(b"test content")
        result = upload_service.validate_file(file, "test.xlsx")
        assert result.is_valid
        assert result.file_extension == ".xlsx"
    
    def test_valid_xls_file(self, upload_service):
        """Test that .xls files are accepted."""
        file = io.BytesIO(b"test content")
        result = upload_service.validate_file(file, "test.xls")
        assert result.is_valid
        assert result.file_extension == ".xls"
    
    def test_valid_csv_file(self, upload_service):
        """Test that .csv files are accepted."""
        file = io.BytesIO(b"test content")
        result = upload_service.validate_file(file, "test.csv")
        assert result.is_valid
        assert result.file_extension == ".csv"
    
    def test_invalid_txt_file(self, upload_service):
        """Test that .txt files are rejected."""
        file = io.BytesIO(b"test content")
        result = upload_service.validate_file(file, "test.txt")
        assert not result.is_valid
        assert "Invalid file type" in result.error_message
    
    def test_case_insensitive_extension(self, upload_service):
        """Test that file extension validation is case-insensitive."""
        file = io.BytesIO(b"test content")
        result = upload_service.validate_file(file, "test.XLSX")
        assert result.is_valid
    
    def test_file_size_at_limit(self, upload_service):
        """Test file exactly at 50MB limit is accepted."""
        file_size_bytes = int(50 * 1024 * 1024)
        file = io.BytesIO(b"x" * file_size_bytes)
        result = upload_service.validate_file(file, "test.xlsx")
        assert result.is_valid
    
    def test_file_size_just_over_limit(self, upload_service):
        """Test file just over 50MB limit is rejected."""
        file_size_bytes = int(50.1 * 1024 * 1024)
        file = io.BytesIO(b"x" * file_size_bytes)
        result = upload_service.validate_file(file, "test.xlsx")
        assert not result.is_valid
        assert "exceeds" in result.error_message.lower()


class TestUploadServiceReadFile:
    """Unit tests for UploadService read_file method."""
    
    def test_read_csv_file(self, upload_service):
        """Test reading a valid CSV file."""
        csv_content = b"col1,col2,col3\n1,2,3\n4,5,6"
        file = io.BytesIO(csv_content)
        df = upload_service.read_file(file, "test.csv")
        assert len(df) == 2
        assert list(df.columns) == ["col1", "col2", "col3"]
    
    def test_read_empty_csv_raises_error(self, upload_service):
        """Test that reading an empty CSV raises ValueError."""
        file = io.BytesIO(b"")
        with pytest.raises(ValueError, match="empty|no data"):
            upload_service.read_file(file, "test.csv")
    
    def test_unsupported_format_raises_error(self, upload_service):
        """Test that unsupported format raises ValueError."""
        file = io.BytesIO(b"test content")
        with pytest.raises(ValueError, match="Unsupported file format"):
            upload_service.read_file(file, "test.txt")


class TestUploadServicePreview:
    """Unit tests for UploadService get_preview method."""
    
    def test_preview_returns_first_n_rows(self, upload_service):
        """Test that preview returns the first N rows."""
        import pandas as pd
        df = pd.DataFrame({"col1": range(20)})
        preview = upload_service.get_preview(df, rows=5)
        assert len(preview) == 5
        assert list(preview["col1"]) == [0, 1, 2, 3, 4]
    
    def test_preview_default_10_rows(self, upload_service):
        """Test that default preview is 10 rows."""
        import pandas as pd
        df = pd.DataFrame({"col1": range(20)})
        preview = upload_service.get_preview(df)
        assert len(preview) == 10
    
    def test_preview_with_fewer_rows_than_requested(self, upload_service):
        """Test preview when DataFrame has fewer rows than requested."""
        import pandas as pd
        df = pd.DataFrame({"col1": range(3)})
        preview = upload_service.get_preview(df, rows=10)
        assert len(preview) == 3
