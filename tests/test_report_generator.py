"""
Property-based tests for the Report Generator.

Tests Properties 22, 23, and 27 from the design document:
- Property 22: Excel Export Round-Trip
- Property 23: CSV Export Round-Trip
- Property 27: Audit Log Completeness
"""

import os
import tempfile
from datetime import datetime

import pandas as pd
import pytest
from hypothesis import given, settings, strategies as st

from src.models import AggregatedAccount, ProcessingStats
from src.report_generator import ReportGenerator


# =============================================================================
# Hypothesis Strategies for Report Generator Tests
# =============================================================================

# Strategy for generating valid account numbers (9-18 digits)
valid_account_numbers = st.text(
    alphabet="0123456789",
    min_size=9,
    max_size=18
)

# Strategy for generating valid IFSC codes
valid_ifsc_codes = st.text(
    alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789",
    min_size=11,
    max_size=11
)

# Strategy for generating valid amounts
valid_amounts = st.floats(
    min_value=0.01, 
    max_value=1e10, 
    allow_nan=False, 
    allow_infinity=False
)

# Strategy for generating bank names (simple strings without special chars)
bank_names = st.text(
    alphabet=st.characters(whitelist_categories=('L', 'N', 'Zs')),
    min_size=1,
    max_size=50
).filter(lambda x: x.strip() != '')

# Strategy for generating addresses
addresses = st.text(
    alphabet=st.characters(whitelist_categories=('L', 'N', 'Zs', 'P')),
    min_size=1,
    max_size=100
).filter(lambda x: x.strip() != '')


# Strategy for generating acknowledgement numbers (semicolon-separated)
ack_numbers = st.lists(
    st.text(alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", min_size=5, max_size=15),
    min_size=1,
    max_size=5
).map(lambda x: ";".join(x))

# Strategy for generating risk scores
risk_scores = st.floats(min_value=0.0, max_value=100.0, allow_nan=False, allow_infinity=False)

# Strategy for generating AggregatedAccount objects
aggregated_account_strategy = st.builds(
    AggregatedAccount,
    account_number=valid_account_numbers,
    bank_name=bank_names,
    ifsc_code=valid_ifsc_codes,
    address=addresses,
    total_transactions=st.integers(min_value=1, max_value=1000),
    acknowledgement_numbers=ack_numbers,
    total_amount=valid_amounts,
    total_disputed_amount=valid_amounts,
    risk_score=risk_scores
)

# Strategy for generating lists of AggregatedAccount objects
aggregated_accounts_list = st.lists(
    aggregated_account_strategy,
    min_size=1,
    max_size=20
)

# Strategy for generating filenames
filenames = st.text(
    alphabet="abcdefghijklmnopqrstuvwxyz0123456789_-.",
    min_size=5,
    max_size=50
).filter(lambda x: x.strip() != '' and not x.startswith('.'))

# Strategy for generating error messages
error_messages = st.lists(
    st.text(min_size=5, max_size=100).filter(lambda x: x.strip() != ''),
    min_size=0,
    max_size=10
)


# =============================================================================
# Property Tests
# =============================================================================

class TestReportGeneratorProperties:
    """Property-based tests for ReportGenerator."""

    
    # Feature: fraud-analysis-app, Property 22: Excel Export Round-Trip
    # Validates: Requirements 5.1, 5.4
    @given(accounts=aggregated_accounts_list)
    @settings(max_examples=100, deadline=None)
    def test_excel_export_round_trip(self, accounts: list):
        """
        Property 22: Excel Export Round-Trip
        
        For any list of aggregated accounts, exporting to Excel and then 
        reading back should produce equivalent data (same account numbers, 
        amounts, and transaction counts).
        """
        generator = ReportGenerator()
        
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            filepath = tmp.name
        
        try:
            # Export to Excel
            generator.generate_excel(accounts, filepath)
            
            # Read back from Excel, treating account number as string
            df = pd.read_excel(
                filepath, 
                engine='openpyxl',
                dtype={"Fraudster Bank Account Number": str}
            )
            
            # Verify row count matches
            assert len(df) == len(accounts), \
                f"Row count mismatch: expected {len(accounts)}, got {len(df)}"
            
            # Verify each account's key fields
            for i, account in enumerate(accounts):
                row = df.iloc[i]
                
                # Check account number
                assert str(row["Fraudster Bank Account Number"]) == account.account_number, \
                    f"Account number mismatch at row {i}"
                
                # Check total amount (with floating point tolerance)
                assert abs(float(row["Total Amount"]) - account.total_amount) < 0.01, \
                    f"Total amount mismatch at row {i}"
                
                # Check transaction count
                assert int(row["Total Transactions"]) == account.total_transactions, \
                    f"Transaction count mismatch at row {i}"
        finally:
            # Clean up temp file
            if os.path.exists(filepath):
                os.unlink(filepath)

    
    # Feature: fraud-analysis-app, Property 23: CSV Export Round-Trip
    # Validates: Requirements 5.5
    @given(accounts=aggregated_accounts_list)
    @settings(max_examples=100, deadline=None)
    def test_csv_export_round_trip(self, accounts: list):
        """
        Property 23: CSV Export Round-Trip
        
        For any list of aggregated accounts, exporting to CSV and then 
        reading back should produce equivalent data.
        """
        generator = ReportGenerator()
        
        with tempfile.NamedTemporaryFile(suffix='.csv', delete=False, mode='w') as tmp:
            filepath = tmp.name
        
        try:
            # Export to CSV
            generator.generate_csv(accounts, filepath)
            
            # Read back from CSV, treating account number as string
            df = pd.read_csv(filepath, dtype={"Fraudster Bank Account Number": str})
            
            # Verify row count matches
            assert len(df) == len(accounts), \
                f"Row count mismatch: expected {len(accounts)}, got {len(df)}"
            
            # Verify each account's key fields
            for i, account in enumerate(accounts):
                row = df.iloc[i]
                
                # Check account number (CSV may read as int, so convert to string)
                assert str(row["Fraudster Bank Account Number"]) == account.account_number, \
                    f"Account number mismatch at row {i}"
                
                # Check total amount (with floating point tolerance)
                assert abs(float(row["Total Amount"]) - account.total_amount) < 0.01, \
                    f"Total amount mismatch at row {i}"
                
                # Check transaction count
                assert int(row["Total Transactions"]) == account.total_transactions, \
                    f"Transaction count mismatch at row {i}"
        finally:
            # Clean up temp file
            if os.path.exists(filepath):
                os.unlink(filepath)

    
    # Feature: fraud-analysis-app, Property 27: Audit Log Completeness
    # Validates: Requirements 9.1, 9.2, 9.3, 9.4, 9.5
    @given(
        filename=filenames,
        rows_processed=st.integers(min_value=0, max_value=100000),
        errors=error_messages
    )
    @settings(max_examples=100, deadline=None)
    def test_audit_log_completeness(
        self, 
        filename: str, 
        rows_processed: int, 
        errors: list
    ):
        """
        Property 27: Audit Log Completeness
        
        For any processing session, the generated audit log should contain:
        timestamp, input filename, rows processed count, and any errors encountered.
        """
        generator = ReportGenerator()
        timestamp = datetime.now()
        
        # Generate audit log
        audit_log = generator.generate_audit_log(
            input_filename=filename,
            rows_processed=rows_processed,
            errors_encountered=errors,
            timestamp=timestamp
        )
        
        # Verify timestamp is present
        timestamp_str = timestamp.strftime('%Y-%m-%d %H:%M:%S')
        assert timestamp_str in audit_log, \
            f"Timestamp '{timestamp_str}' not found in audit log"
        
        # Verify filename is present
        assert filename in audit_log, \
            f"Filename '{filename}' not found in audit log"
        
        # Verify rows processed count is present
        assert str(rows_processed) in audit_log, \
            f"Rows processed count '{rows_processed}' not found in audit log"
        
        # Verify all errors are present
        for error in errors:
            assert error in audit_log, \
                f"Error '{error}' not found in audit log"


# =============================================================================
# Unit Tests for Edge Cases
# =============================================================================

class TestReportGeneratorEdgeCases:
    """Unit tests for edge cases in ReportGenerator."""
    
    def test_empty_accounts_list_excel(self):
        """Test Excel generation with empty accounts list."""
        generator = ReportGenerator()
        
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            filepath = tmp.name
        
        try:
            generator.generate_excel([], filepath)
            df = pd.read_excel(filepath, engine='openpyxl')
            assert len(df) == 0
            assert list(df.columns) == generator.OUTPUT_COLUMNS
        finally:
            if os.path.exists(filepath):
                os.unlink(filepath)
    
    def test_empty_accounts_list_csv(self):
        """Test CSV generation with empty accounts list."""
        generator = ReportGenerator()
        
        with tempfile.NamedTemporaryFile(suffix='.csv', delete=False, mode='w') as tmp:
            filepath = tmp.name
        
        try:
            generator.generate_csv([], filepath)
            df = pd.read_csv(filepath)
            assert len(df) == 0
            assert list(df.columns) == generator.OUTPUT_COLUMNS
        finally:
            if os.path.exists(filepath):
                os.unlink(filepath)
    
    def test_audit_log_no_errors(self):
        """Test audit log generation with no errors."""
        generator = ReportGenerator()
        
        audit_log = generator.generate_audit_log(
            input_filename="test.xlsx",
            rows_processed=100,
            errors_encountered=[]
        )
        
        assert "test.xlsx" in audit_log
        assert "100" in audit_log
        assert "No errors encountered" in audit_log
    
    def test_generate_excel_bytes(self):
        """Test Excel bytes generation."""
        generator = ReportGenerator()
        account = AggregatedAccount(
            account_number="123456789012",
            bank_name="Test Bank",
            ifsc_code="TEST0001234",
            address="Test Address",
            district="Test District",
            state="Test State",
            total_transactions=5,
            acknowledgement_numbers="ACK1;ACK2",
            total_amount=50000.0,
            total_disputed_amount=50000.0,
            risk_score=75.0
        )
        
        excel_bytes = generator.generate_excel_bytes([account])
        assert isinstance(excel_bytes, bytes)
        assert len(excel_bytes) > 0
    
    def test_generate_csv_bytes(self):
        """Test CSV bytes generation."""
        generator = ReportGenerator()
        account = AggregatedAccount(
            account_number="123456789012",
            bank_name="Test Bank",
            ifsc_code="TEST0001234",
            address="Test Address",
            district="Test District",
            state="Test State",
            total_transactions=5,
            acknowledgement_numbers="ACK1;ACK2",
            total_amount=50000.0,
            total_disputed_amount=50000.0,
            risk_score=75.0
        )
        
        csv_bytes = generator.generate_csv_bytes([account])
        assert isinstance(csv_bytes, bytes)
        assert b"123456789012" in csv_bytes
        assert b"Test Bank" in csv_bytes

    
    def test_generate_pdf(self):
        """Test PDF generation with sample data."""
        generator = ReportGenerator()
        account = AggregatedAccount(
            account_number="123456789012",
            bank_name="Test Bank",
            ifsc_code="TEST0001234",
            address="Test Address",
            district="Test District",
            state="Test State",
            total_transactions=5,
            acknowledgement_numbers="ACK1;ACK2",
            total_amount=50000.0,
            total_disputed_amount=50000.0,
            risk_score=75.0
        )
        
        stats = ProcessingStats(
            total_input_rows=100,
            rows_processed=95,
            rows_with_errors=5,
            unique_accounts=10,
            total_fraud_amount=500000.0,
            total_disputed_amount=500000.0,
            average_amount_per_account=50000.0,
            top_accounts_by_amount=[account],
            processing_timestamp=datetime.now(),
            input_filename="test.xlsx"
        )
        
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            filepath = tmp.name
        
        try:
            generator.generate_pdf([account], stats, filepath)
            # Verify file was created and has content
            assert os.path.exists(filepath)
            assert os.path.getsize(filepath) > 0
        finally:
            if os.path.exists(filepath):
                os.unlink(filepath)
    
    def test_generate_pdf_bytes(self):
        """Test PDF bytes generation."""
        generator = ReportGenerator()
        account = AggregatedAccount(
            account_number="123456789012",
            bank_name="Test Bank",
            ifsc_code="TEST0001234",
            address="Test Address",
            district="Test District",
            state="Test State",
            total_transactions=5,
            acknowledgement_numbers="ACK1;ACK2",
            total_amount=50000.0,
            total_disputed_amount=50000.0,
            risk_score=75.0
        )
        
        stats = ProcessingStats(
            total_input_rows=100,
            rows_processed=95,
            rows_with_errors=5,
            unique_accounts=10,
            total_fraud_amount=500000.0,
            total_disputed_amount=500000.0,
            average_amount_per_account=50000.0,
            top_accounts_by_amount=[account],
            processing_timestamp=datetime.now(),
            input_filename="test.xlsx"
        )
        
        pdf_bytes = generator.generate_pdf_bytes([account], stats)
        assert isinstance(pdf_bytes, bytes)
        assert len(pdf_bytes) > 0
        # PDF files start with %PDF
        assert pdf_bytes[:4] == b'%PDF'
    
    def test_generate_pdf_with_quality_metrics(self):
        """Test PDF generation with quality metrics."""
        generator = ReportGenerator()
        account = AggregatedAccount(
            account_number="123456789012",
            bank_name="Test Bank",
            ifsc_code="TEST0001234",
            address="Test Address",
            district="Test District",
            state="Test State",
            total_transactions=5,
            acknowledgement_numbers="ACK1;ACK2",
            total_amount=50000.0,
            total_disputed_amount=50000.0,
            risk_score=75.0
        )
        
        stats = ProcessingStats(
            total_input_rows=100,
            rows_processed=95,
            rows_with_errors=5,
            unique_accounts=10,
            total_fraud_amount=500000.0,
            total_disputed_amount=500000.0,
            average_amount_per_account=50000.0,
            top_accounts_by_amount=[account],
            processing_timestamp=datetime.now(),
            input_filename="test.xlsx"
        )
        
        quality_metrics = {
            "completeness": 0.95,
            "valid_accounts": 90,
            "invalid_accounts": 10
        }
        
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            filepath = tmp.name
        
        try:
            generator.generate_pdf([account], stats, filepath, quality_metrics)
            assert os.path.exists(filepath)
            assert os.path.getsize(filepath) > 0
        finally:
            if os.path.exists(filepath):
                os.unlink(filepath)
    
    def test_generate_pdf_empty_accounts(self):
        """Test PDF generation with empty accounts list."""
        generator = ReportGenerator()
        
        stats = ProcessingStats(
            total_input_rows=0,
            rows_processed=0,
            rows_with_errors=0,
            unique_accounts=0,
            total_fraud_amount=0.0,
            total_disputed_amount=0.0,
            average_amount_per_account=0.0,
            top_accounts_by_amount=[],
            processing_timestamp=datetime.now(),
            input_filename="empty.xlsx"
        )
        
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            filepath = tmp.name
        
        try:
            generator.generate_pdf([], stats, filepath)
            assert os.path.exists(filepath)
            assert os.path.getsize(filepath) > 0
        finally:
            if os.path.exists(filepath):
                os.unlink(filepath)
