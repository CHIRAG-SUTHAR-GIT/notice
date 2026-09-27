"""
Integration tests for the Fraud Analysis Application.

Tests the complete end-to-end flow: upload → process → download
with various file formats, column naming conventions, and data sizes.
"""

import io
import os
import tempfile
from datetime import datetime
from typing import List

import pandas as pd
import pytest

from src.upload_service import UploadService
from src.column_detector import ColumnDetector
from src.data_processor import DataProcessor
from src.validation_engine import ValidationEngine
from src.aggregation_engine import AggregationEngine
from src.report_generator import ReportGenerator
from src.dashboard import Dashboard
from src.models import ColumnMapping, AggregatedAccount


# =============================================================================
# Test Fixtures
# =============================================================================

@pytest.fixture
def upload_service():
    """Fixture providing an UploadService instance."""
    return UploadService()


@pytest.fixture
def column_detector():
    """Fixture providing a ColumnDetector instance."""
    return ColumnDetector()


@pytest.fixture
def data_processor():
    """Fixture providing a DataProcessor instance."""
    return DataProcessor()


@pytest.fixture
def validation_engine():
    """Fixture providing a ValidationEngine instance."""
    return ValidationEngine()


@pytest.fixture
def aggregation_engine():
    """Fixture providing an AggregationEngine instance."""
    return AggregationEngine()


@pytest.fixture
def report_generator():
    """Fixture providing a ReportGenerator instance."""
    return ReportGenerator()


@pytest.fixture
def dashboard():
    """Fixture providing a Dashboard instance."""
    return Dashboard()


def create_test_dataframe(num_rows: int, num_accounts: int = None) -> pd.DataFrame:
    """
    Create a test DataFrame with fraud transaction data.
    
    Args:
        num_rows: Number of rows to generate.
        num_accounts: Number of unique accounts (defaults to num_rows // 3).
        
    Returns:
        DataFrame with test transaction data.
    """
    if num_accounts is None:
        num_accounts = max(1, num_rows // 3)
    
    # Generate account numbers
    accounts = [f"{100000000 + i:012d}" for i in range(num_accounts)]
    
    data = []
    for i in range(num_rows):
        account_idx = i % num_accounts
        data.append({
            "Sr No": i + 1,
            "Ack No": f"ACK{i + 1:06d}",
            "Bank Account No": accounts[account_idx],
            "IFSC Code": f"SBIN{account_idx:07d}",
            "Address": f"Address {account_idx + 1}, City {account_idx % 10}",
            "Amount": 1000.0 + (i * 100),
            "Disputed Amount": 500.0 + (i * 50),
            "Bank Name": ["State Bank of India", "HDFC Bank", "ICICI Bank", 
                         "Axis Bank", "Punjab National Bank"][account_idx % 5]
        })
    
    return pd.DataFrame(data)


def create_csv_bytes(df: pd.DataFrame) -> bytes:
    """Convert DataFrame to CSV bytes."""
    buffer = io.StringIO()
    df.to_csv(buffer, index=False)
    return buffer.getvalue().encode('utf-8')


def create_excel_bytes(df: pd.DataFrame) -> bytes:
    """Convert DataFrame to Excel bytes."""
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)
    return buffer.getvalue()


# =============================================================================
# End-to-End Integration Tests
# =============================================================================

class TestEndToEndProcessing:
    """End-to-end tests for the complete processing pipeline."""
    
    def test_complete_pipeline_csv(
        self, upload_service, column_detector, data_processor,
        validation_engine, aggregation_engine, report_generator, dashboard
    ):
        """
        Test complete pipeline: upload CSV → detect columns → process → 
        validate → aggregate → generate reports.
        
        Requirements: All
        """
        # Step 1: Create test data
        df_original = create_test_dataframe(num_rows=10, num_accounts=3)
        csv_bytes = create_csv_bytes(df_original)
        file = io.BytesIO(csv_bytes)
        filename = "test_transactions.csv"
        
        # Step 2: Upload and validate file
        validation_result = upload_service.validate_file(file, filename)
        assert validation_result.is_valid, f"File validation failed: {validation_result.error_message}"
        
        # Step 3: Read file
        file.seek(0)
        df = upload_service.read_file(file, filename)
        assert len(df) == 10, "Should have 10 rows"
        
        # Step 4: Detect columns
        headers = list(df.columns)
        mapping = column_detector.detect_columns(headers)
        assert mapping.bank_account_number is not None, "Should detect bank account column"
        assert mapping.amount is not None, "Should detect amount column"
        
        # Step 5: Clean data
        df_cleaned = data_processor.clean_dataframe(df, mapping)
        assert len(df_cleaned) == 10, "Should still have 10 rows after cleaning"
        
        # Step 6: Validate data
        validation = validation_engine.validate_dataframe(df_cleaned, mapping)
        assert validation.is_valid, f"Validation failed: {validation.critical_errors}"
        
        # Step 7: Aggregate by account
        accounts = aggregation_engine.aggregate_by_account(df_cleaned, mapping)
        assert len(accounts) == 3, "Should have 3 unique accounts"
        
        # Step 8: Sort results
        sorted_accounts = aggregation_engine.sort_results(accounts)
        assert len(sorted_accounts) == 3
        
        # Verify sorting order (descending by amount)
        for i in range(len(sorted_accounts) - 1):
            assert sorted_accounts[i].total_amount >= sorted_accounts[i + 1].total_amount
        
        # Step 9: Calculate statistics
        stats = dashboard.calculate_statistics(
            sorted_accounts, 
            total_input_rows=10,
            input_filename=filename
        )
        assert stats.unique_accounts == 3
        assert stats.total_fraud_amount > 0
        
        # Step 10: Generate reports
        with tempfile.TemporaryDirectory() as tmpdir:
            # Excel report
            excel_path = os.path.join(tmpdir, "output.xlsx")
            report_generator.generate_excel(sorted_accounts, excel_path)
            assert os.path.exists(excel_path)
            
            # CSV report
            csv_path = os.path.join(tmpdir, "output.csv")
            report_generator.generate_csv(sorted_accounts, csv_path)
            assert os.path.exists(csv_path)
            
            # Verify Excel round-trip
            df_excel = pd.read_excel(excel_path)
            assert len(df_excel) == 3
            
            # Verify CSV round-trip
            df_csv = pd.read_csv(csv_path)
            assert len(df_csv) == 3
    
    def test_complete_pipeline_excel(
        self, upload_service, column_detector, data_processor,
        validation_engine, aggregation_engine, report_generator
    ):
        """
        Test complete pipeline with Excel file input.
        
        Requirements: All
        """
        # Create test data
        df_original = create_test_dataframe(num_rows=20, num_accounts=5)
        excel_bytes = create_excel_bytes(df_original)
        file = io.BytesIO(excel_bytes)
        filename = "test_transactions.xlsx"
        
        # Upload and validate
        validation_result = upload_service.validate_file(file, filename)
        assert validation_result.is_valid
        
        # Read file
        file.seek(0)
        df = upload_service.read_file(file, filename)
        assert len(df) == 20
        
        # Detect columns
        mapping = column_detector.detect_columns(list(df.columns))
        assert mapping.bank_account_number is not None
        
        # Process pipeline
        df_cleaned = data_processor.clean_dataframe(df, mapping)
        validation = validation_engine.validate_dataframe(df_cleaned, mapping)
        assert validation.is_valid
        
        accounts = aggregation_engine.aggregate_by_account(df_cleaned, mapping)
        assert len(accounts) == 5


# =============================================================================
# Column Naming Convention Tests
# =============================================================================

class TestColumnNamingConventions:
    """Tests for various column naming conventions."""
    
    @pytest.mark.parametrize("account_col_name", [
        "Bank Account No",
        "bank account no",
        "BANK ACCOUNT NO",
        "Bank A/C No",
        "Account Number",
        "account number",
        "Beneficiary Account",
        "beneficiary ac",
        "ac no",
        "A/C No"
    ])
    def test_account_column_variants(
        self, column_detector, data_processor, validation_engine, aggregation_engine,
        account_col_name
    ):
        """
        Test that various account column naming conventions are detected.
        
        Requirements: 2.8
        """
        df = pd.DataFrame({
            account_col_name: ["123456789012", "987654321098"],
            "Amount": [1000.0, 2000.0]
        })
        
        mapping = column_detector.detect_columns(list(df.columns))
        assert mapping.bank_account_number == account_col_name, \
            f"Failed to detect '{account_col_name}' as bank account column"
    
    @pytest.mark.parametrize("amount_col_name", [
        "Amount",
        "amount",
        "AMOUNT",
        "Transaction Amount",
        "txn amount",
        "Transfer Amount",
        "Fraud Amount"
    ])
    def test_amount_column_variants(self, column_detector, amount_col_name):
        """
        Test that various amount column naming conventions are detected.
        
        Requirements: 2.11
        """
        df = pd.DataFrame({
            "Bank Account No": ["123456789012"],
            amount_col_name: [1000.0]
        })
        
        mapping = column_detector.detect_columns(list(df.columns))
        assert mapping.amount == amount_col_name, \
            f"Failed to detect '{amount_col_name}' as amount column"
    
    @pytest.mark.parametrize("ifsc_col_name", [
        "IFSC Code",
        "ifsc code",
        "IFSC",
        "ifsc",
        "Bank Code"
    ])
    def test_ifsc_column_variants(self, column_detector, ifsc_col_name):
        """
        Test that various IFSC column naming conventions are detected.
        
        Requirements: 2.9
        """
        df = pd.DataFrame({
            "Bank Account No": ["123456789012"],
            "Amount": [1000.0],
            ifsc_col_name: ["SBIN0001234"]
        })
        
        mapping = column_detector.detect_columns(list(df.columns))
        assert mapping.ifsc_code == ifsc_col_name, \
            f"Failed to detect '{ifsc_col_name}' as IFSC column"
    
    def test_mixed_case_and_whitespace_headers(
        self, column_detector, data_processor, validation_engine, aggregation_engine
    ):
        """
        Test processing with mixed case and whitespace in headers.
        
        Requirements: 2.2
        """
        df = pd.DataFrame({
            "  Bank Account No  ": ["123456789012", "987654321098"],
            "  AMOUNT  ": [1000.0, 2000.0],
            "  ifsc CODE  ": ["SBIN0001234", "HDFC0005678"],
            "  Bank Name  ": ["SBI", "HDFC"]
        })
        
        # Strip column names for detection
        df.columns = [col.strip() for col in df.columns]
        
        mapping = column_detector.detect_columns(list(df.columns))
        assert mapping.bank_account_number is not None
        assert mapping.amount is not None


# =============================================================================
# Data Size Tests
# =============================================================================

class TestDataSizes:
    """Tests for various data sizes."""
    
    def test_10_rows(
        self, upload_service, column_detector, data_processor,
        validation_engine, aggregation_engine
    ):
        """Test processing with 10 rows."""
        df = create_test_dataframe(num_rows=10, num_accounts=3)
        self._run_pipeline(df, column_detector, data_processor, 
                          validation_engine, aggregation_engine, 
                          expected_accounts=3)
    
    def test_100_rows(
        self, upload_service, column_detector, data_processor,
        validation_engine, aggregation_engine
    ):
        """Test processing with 100 rows."""
        df = create_test_dataframe(num_rows=100, num_accounts=20)
        self._run_pipeline(df, column_detector, data_processor,
                          validation_engine, aggregation_engine,
                          expected_accounts=20)
    
    def test_1000_rows(
        self, upload_service, column_detector, data_processor,
        validation_engine, aggregation_engine
    ):
        """Test processing with 1000 rows."""
        df = create_test_dataframe(num_rows=1000, num_accounts=100)
        self._run_pipeline(df, column_detector, data_processor,
                          validation_engine, aggregation_engine,
                          expected_accounts=100)
    
    @pytest.mark.slow
    def test_10000_rows(
        self, upload_service, column_detector, data_processor,
        validation_engine, aggregation_engine
    ):
        """Test processing with 10000+ rows."""
        df = create_test_dataframe(num_rows=10000, num_accounts=500)
        self._run_pipeline(df, column_detector, data_processor,
                          validation_engine, aggregation_engine,
                          expected_accounts=500)
    
    def _run_pipeline(
        self, df, column_detector, data_processor,
        validation_engine, aggregation_engine, expected_accounts
    ):
        """Helper to run the processing pipeline."""
        mapping = column_detector.detect_columns(list(df.columns))
        assert mapping.bank_account_number is not None
        assert mapping.amount is not None
        
        df_cleaned = data_processor.clean_dataframe(df, mapping)
        validation = validation_engine.validate_dataframe(df_cleaned, mapping)
        assert validation.is_valid
        
        accounts = aggregation_engine.aggregate_by_account(df_cleaned, mapping)
        assert len(accounts) == expected_accounts


# =============================================================================
# Missing Data Scenarios
# =============================================================================

class TestMissingDataScenarios:
    """Tests for handling missing data."""
    
    def test_missing_optional_columns(
        self, column_detector, data_processor, validation_engine, aggregation_engine
    ):
        """
        Test processing when optional columns are missing.
        
        Requirements: 10.3, 10.4
        """
        # Only required columns present
        df = pd.DataFrame({
            "Bank Account No": ["123456789012", "987654321098", "123456789012"],
            "Amount": [1000.0, 2000.0, 1500.0]
        })
        
        mapping = column_detector.detect_columns(list(df.columns))
        assert mapping.bank_account_number is not None
        assert mapping.amount is not None
        assert mapping.ifsc_code is None  # Optional, not present
        
        df_cleaned = data_processor.clean_dataframe(df, mapping)
        validation = validation_engine.validate_dataframe(df_cleaned, mapping)
        
        # Should be valid even without optional columns
        assert validation.is_valid
        
        accounts = aggregation_engine.aggregate_by_account(df_cleaned, mapping)
        assert len(accounts) == 2
    
    def test_some_rows_missing_ifsc(
        self, column_detector, data_processor, validation_engine, aggregation_engine
    ):
        """
        Test processing when some rows have missing IFSC codes.
        
        Requirements: 10.3
        """
        df = pd.DataFrame({
            "Bank Account No": ["123456789012", "987654321098", "555555555555"],
            "Amount": [1000.0, 2000.0, 3000.0],
            "IFSC Code": ["SBIN0001234", None, "HDFC0005678"]
        })
        
        mapping = column_detector.detect_columns(list(df.columns))
        df_cleaned = data_processor.clean_dataframe(df, mapping)
        validation = validation_engine.validate_dataframe(df_cleaned, mapping)
        
        # Should be valid but with warnings
        assert validation.is_valid
        assert len(validation.warnings) > 0  # Should have warning for missing IFSC
    
    def test_some_rows_missing_address(
        self, column_detector, data_processor, validation_engine, aggregation_engine
    ):
        """
        Test processing when some rows have missing addresses.
        
        Requirements: 10.4
        """
        df = pd.DataFrame({
            "Bank Account No": ["123456789012", "987654321098"],
            "Amount": [1000.0, 2000.0],
            "Address": ["123 Main St", None]
        })
        
        mapping = column_detector.detect_columns(list(df.columns))
        df_cleaned = data_processor.clean_dataframe(df, mapping)
        validation = validation_engine.validate_dataframe(df_cleaned, mapping)
        
        assert validation.is_valid
        assert len(validation.warnings) > 0  # Should have warning for missing address
    
    def test_invalid_amount_format(
        self, column_detector, data_processor, validation_engine
    ):
        """
        Test processing when amount has invalid format.
        
        Requirements: 10.5
        """
        df = pd.DataFrame({
            "Bank Account No": ["123456789012", "987654321098"],
            "Amount": ["invalid", "2000.0"]
        })
        
        mapping = column_detector.detect_columns(list(df.columns))
        df_cleaned = data_processor.clean_dataframe(df, mapping)
        
        # Invalid amount should be converted to 0
        assert df_cleaned[mapping.amount].iloc[0] == 0.0
        assert df_cleaned[mapping.amount].iloc[1] == 2000.0
    
    def test_empty_rows_removed(
        self, column_detector, data_processor, validation_engine
    ):
        """
        Test that completely empty rows are removed.
        
        Requirements: 3.1
        """
        df = pd.DataFrame({
            "Bank Account No": ["123456789012", "", "987654321098", None],
            "Amount": [1000.0, None, 2000.0, None],
            "Bank Name": ["SBI", "", "HDFC", ""]
        })
        
        mapping = column_detector.detect_columns(list(df.columns))
        df_cleaned = data_processor.clean_dataframe(df, mapping)
        
        # Empty rows should be removed
        assert len(df_cleaned) < 4
    
    def test_duplicate_acknowledgement_numbers(
        self, column_detector, data_processor, validation_engine
    ):
        """
        Test detection of duplicate acknowledgement numbers.
        
        Requirements: 3.9
        """
        df = pd.DataFrame({
            "Bank Account No": ["123456789012", "987654321098", "555555555555"],
            "Amount": [1000.0, 2000.0, 3000.0],
            "Ack No": ["ACK001", "ACK001", "ACK002"]  # Duplicate ACK001
        })
        
        mapping = column_detector.detect_columns(list(df.columns))
        df_cleaned = data_processor.clean_dataframe(df, mapping)
        validation = validation_engine.validate_dataframe(df_cleaned, mapping)
        
        # Should have warning about duplicate
        assert any("ACK001" in w for w in validation.warnings)


# =============================================================================
# Multi-Sheet Excel Tests
# =============================================================================

class TestMultiSheetExcel:
    """Tests for multi-sheet Excel files."""
    
    def test_uses_first_sheet(self, upload_service, column_detector):
        """
        Test that multi-sheet Excel files use the first sheet.
        
        Requirements: All (multi-sheet handling)
        """
        # Create multi-sheet Excel file
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            # First sheet with valid data
            df1 = pd.DataFrame({
                "Bank Account No": ["123456789012", "987654321098"],
                "Amount": [1000.0, 2000.0]
            })
            df1.to_excel(writer, sheet_name='Sheet1', index=False)
            
            # Second sheet with different data
            df2 = pd.DataFrame({
                "Other Column": ["A", "B"],
                "Another Column": [1, 2]
            })
            df2.to_excel(writer, sheet_name='Sheet2', index=False)
        
        buffer.seek(0)
        
        # Read file (should use first sheet)
        df = upload_service.read_file(buffer, "multi_sheet.xlsx")
        
        # Should have columns from first sheet
        assert "Bank Account No" in df.columns
        assert "Amount" in df.columns
        assert len(df) == 2


# =============================================================================
# Report Generation Tests
# =============================================================================

class TestReportGeneration:
    """Tests for report generation."""
    
    def test_excel_report_contains_all_columns(
        self, aggregation_engine, report_generator
    ):
        """Test that Excel report contains all required columns."""
        accounts = [
            AggregatedAccount(
                account_number="123456789012",
                bank_name="State Bank of India",
                ifsc_code="SBIN0001234",
                address="123 Main St",
                district="Mumbai",
                state="Maharashtra",
                total_transactions=5,
                acknowledgement_numbers="ACK001;ACK002;ACK003",
                total_amount=50000.0,
                total_disputed_amount=25000.0,
                risk_score=75.5
            )
        ]
        
        with tempfile.TemporaryDirectory() as tmpdir:
            filepath = os.path.join(tmpdir, "report.xlsx")
            report_generator.generate_excel(accounts, filepath)
            
            df = pd.read_excel(filepath)
            
            # Check all required columns
            expected_columns = [
                "Fraudster Bank Account Number",
                "Bank Name",
                "IFSC Code",
                "Address",
                "District",
                "State",
                "Total Transactions",
                "All Acknowledgement Numbers",
                "Total Amount",
                "Total Disputed Amount",
                "Risk Score"
            ]
            
            for col in expected_columns:
                assert col in df.columns, f"Missing column: {col}"
    
    def test_csv_report_round_trip(self, report_generator):
        """Test CSV export and re-import produces equivalent data."""
        accounts = [
            AggregatedAccount(
                account_number="123456789012",
                bank_name="SBI",
                ifsc_code="SBIN0001234",
                address="Address 1",
                district="District 1",
                state="State 1",
                total_transactions=3,
                acknowledgement_numbers="ACK1;ACK2",
                total_amount=30000.0,
                total_disputed_amount=15000.0,
                risk_score=60.0
            ),
            AggregatedAccount(
                account_number="987654321098",
                bank_name="HDFC",
                ifsc_code="HDFC0005678",
                address="Address 2",
                district="District 2",
                state="State 2",
                total_transactions=2,
                acknowledgement_numbers="ACK3",
                total_amount=20000.0,
                total_disputed_amount=10000.0,
                risk_score=40.0
            )
        ]
        
        with tempfile.TemporaryDirectory() as tmpdir:
            filepath = os.path.join(tmpdir, "report.csv")
            report_generator.generate_csv(accounts, filepath)
            
            df = pd.read_csv(filepath, dtype={"Fraudster Bank Account Number": str})
            
            assert len(df) == 2
            assert df.iloc[0]["Total Amount"] == 30000.0
            assert df.iloc[1]["Total Amount"] == 20000.0
    
    def test_audit_log_generation(self, report_generator):
        """Test audit log contains required information."""
        log = report_generator.generate_audit_log(
            input_filename="test_file.xlsx",
            rows_processed=100,
            errors_encountered=["Error 1", "Error 2"],
            timestamp=datetime(2024, 1, 15, 10, 30, 0)
        )
        
        assert "test_file.xlsx" in log
        assert "100" in log
        assert "Error 1" in log
        assert "Error 2" in log
        assert "2024-01-15" in log


# =============================================================================
# Dashboard and Filter Tests
# =============================================================================

class TestDashboardIntegration:
    """Integration tests for dashboard functionality."""
    
    def test_statistics_calculation(self, dashboard):
        """Test statistics calculation from aggregated accounts."""
        accounts = [
            AggregatedAccount("111", "SBI", "SBIN001", "Addr1", "", "", 5, "ACK1", 50000.0, 25000.0, 70.0),
            AggregatedAccount("222", "HDFC", "HDFC001", "Addr2", "", "", 3, "ACK2", 30000.0, 15000.0, 50.0),
            AggregatedAccount("333", "ICICI", "ICIC001", "Addr3", "", "", 2, "ACK3", 20000.0, 10000.0, 30.0),
        ]
        
        stats = dashboard.calculate_statistics(accounts, total_input_rows=10, input_filename="test.xlsx")
        
        assert stats.unique_accounts == 3
        assert stats.total_fraud_amount == 100000.0
        assert stats.total_disputed_amount == 50000.0
        assert abs(stats.average_amount_per_account - 33333.33) < 1
    
    def test_search_filter(self, dashboard):
        """Test account search functionality."""
        accounts = [
            AggregatedAccount("123456789012", "SBI", "SBIN001", "Addr1", "", "", 5, "ACK1", 50000.0, 25000.0, 70.0),
            AggregatedAccount("987654321098", "HDFC", "HDFC001", "Addr2", "", "", 3, "ACK2", 30000.0, 15000.0, 50.0),
            AggregatedAccount("123000000000", "ICICI", "ICIC001", "Addr3", "", "", 2, "ACK3", 20000.0, 10000.0, 30.0),
        ]
        
        # Search for accounts containing "123"
        results = dashboard.search_accounts(accounts, "123")
        assert len(results) == 2
        
        # Search for specific account
        results = dashboard.search_accounts(accounts, "987654321098")
        assert len(results) == 1
        assert results[0].account_number == "987654321098"
    
    def test_min_transactions_filter(self, dashboard):
        """Test minimum transactions filter."""
        accounts = [
            AggregatedAccount("111", "SBI", "SBIN001", "Addr1", "", "", 5, "ACK1", 50000.0, 25000.0, 70.0),
            AggregatedAccount("222", "HDFC", "HDFC001", "Addr2", "", "", 3, "ACK2", 30000.0, 15000.0, 50.0),
            AggregatedAccount("333", "ICICI", "ICIC001", "Addr3", "", "", 2, "ACK3", 20000.0, 10000.0, 30.0),
        ]
        
        results = dashboard.filter_by_min_transactions(accounts, 3)
        assert len(results) == 2
        
        results = dashboard.filter_by_min_transactions(accounts, 5)
        assert len(results) == 1
    
    def test_min_amount_filter(self, dashboard):
        """Test minimum amount filter."""
        accounts = [
            AggregatedAccount("111", "SBI", "SBIN001", "Addr1", "", "", 5, "ACK1", 50000.0, 25000.0, 70.0),
            AggregatedAccount("222", "HDFC", "HDFC001", "Addr2", "", "", 3, "ACK2", 30000.0, 15000.0, 50.0),
            AggregatedAccount("333", "ICICI", "ICIC001", "Addr3", "", "", 2, "ACK3", 20000.0, 10000.0, 30.0),
        ]
        
        results = dashboard.filter_by_min_amount(accounts, 25000.0)
        assert len(results) == 2
        
        results = dashboard.filter_by_min_amount(accounts, 50000.0)
        assert len(results) == 1
