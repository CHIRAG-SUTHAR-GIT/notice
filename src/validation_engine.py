"""
Validation Engine for the Fraud Analysis Application.

Contains the ValidationEngine class responsible for validating data integrity,
checking field formats, detecting duplicates, and generating quality reports.
"""

import re
from typing import Any, Dict, List, Optional, Set, Tuple

import pandas as pd

from src.models import ColumnMapping, ErrorCategory, ErrorResponse, ValidationResult


# Critical error codes and messages
CRITICAL_ERRORS = {
    "NO_ACCOUNT_COLUMN": "No bank account number column found in the file",
    "NO_AMOUNT_COLUMN": "No amount column found in the file",
    "FILE_CORRUPTED": "File is corrupted or unreadable",
    "FILE_TOO_LARGE": "File exceeds maximum size limit of 50MB",
    "INVALID_FORMAT": "File format is not supported"
}

# Warning error codes and messages
WARNING_ERRORS = {
    "MISSING_IFSC": "IFSC code missing for row {row}",
    "MISSING_ADDRESS": "Address missing for row {row}",
    "INVALID_AMOUNT": "Invalid amount format at row {row}, using 0",
    "DUPLICATE_ACK": "Duplicate acknowledgement number: {ack_no}",
    "INVALID_ACCOUNT": "Invalid account number format at row {row}",
    "INVALID_IFSC_FORMAT": "Invalid IFSC format at row {row}"
}


def _is_missing(value: Any) -> bool:
    """Return True for null, blank, and common spreadsheet null strings."""
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    return text == "" or text.lower() in {"nan", "none", "<na>", "nat"}


class ValidationEngine:
    """
    Validates transaction data for integrity and format compliance.

    Handles:
    - Account number validation (9-18 digits)
    - IFSC code validation (11 alphanumeric characters)
    - Amount validation (positive numbers)
    - Duplicate acknowledgement detection
    - Critical data flagging
    - Data quality report generation
    """

    ACCOUNT_NUMBER_MIN_LENGTH: int = 9
    ACCOUNT_NUMBER_MAX_LENGTH: int = 18
    IFSC_CODE_LENGTH: int = 11

    def validate_account_number(self, account: Optional[str]) -> bool:
        """
        Validate that account number contains 9-18 digits.

        Args:
            account: Account number string to validate

        Returns:
            True if valid (9-18 digits after removing non-digits), False otherwise
        """
        if _is_missing(account):
            return False

        # Extract only digits from the account number
        digits_only = re.sub(r'\D', '', str(account))

        # Check if digit count is within valid range
        return self.ACCOUNT_NUMBER_MIN_LENGTH <= len(digits_only) <= self.ACCOUNT_NUMBER_MAX_LENGTH

    def validate_ifsc_code(self, ifsc: Optional[str]) -> bool:
        """
        Validate that IFSC code is exactly 11 alphanumeric characters.

        Args:
            ifsc: IFSC code string to validate

        Returns:
            True if valid (exactly 11 alphanumeric chars), False otherwise
        """
        if _is_missing(ifsc):
            return False

        ifsc_str = str(ifsc).strip()

        # Check length is exactly 11
        if len(ifsc_str) != self.IFSC_CODE_LENGTH:
            return False

        # Check all characters are alphanumeric
        return ifsc_str.isalnum()

    def validate_amount(self, amount: Any) -> bool:
        """
        Validate that amount is a positive number.

        Args:
            amount: Amount value to validate (can be float, int, or string)

        Returns:
            True if valid (positive number > 0), False otherwise
        """
        if _is_missing(amount):
            return False

        try:
            amount_float = float(amount)
            return amount_float > 0
        except (ValueError, TypeError):
            return False

    def check_duplicate_acknowledgements(
        self, df: pd.DataFrame, ack_col: Optional[str]
    ) -> List[str]:
        """
        Find duplicate acknowledgement numbers in the DataFrame.

        Args:
            df: DataFrame to check
            ack_col: Name of the acknowledgement number column

        Returns:
            List of duplicate acknowledgement numbers
        """
        if ack_col is None or ack_col not in df.columns:
            return []

        # Get non-null acknowledgement numbers
        ack_series = df[ack_col].dropna()

        # Convert to string for consistent comparison
        ack_series = ack_series.astype(str)

        # Filter out empty strings and null-like text values.
        ack_series = ack_series[~ack_series.str.strip().str.lower().isin(['', 'nan', 'none', '<na>', 'nat'])]

        # Find duplicates
        duplicates = ack_series[ack_series.duplicated(keep=False)].unique().tolist()

        return duplicates

    def _check_critical_missing_data(
        self, row: pd.Series, mapping: ColumnMapping, row_idx: int
    ) -> Tuple[bool, Optional[ErrorResponse]]:
        """
        Check if a row has critical missing data (account number or amount).

        Args:
            row: DataFrame row to check
            mapping: Column mapping
            row_idx: Row index for error reporting

        Returns:
            Tuple of (has_critical_missing, error_response)
        """
        # Check account number
        if mapping.bank_account_number:
            account_val = row.get(mapping.bank_account_number)
            if _is_missing(account_val):
                return True, ErrorResponse(
                    category=ErrorCategory.CRITICAL,
                    code="MISSING_ACCOUNT",
                    message=f"Missing bank account number at row {row_idx}",
                    row_number=row_idx,
                    field_name="bank_account_number"
                )

        # Check amount
        if mapping.amount:
            amount_val = row.get(mapping.amount)
            if _is_missing(amount_val):
                return True, ErrorResponse(
                    category=ErrorCategory.CRITICAL,
                    code="MISSING_AMOUNT",
                    message=f"Missing amount at row {row_idx}",
                    row_number=row_idx,
                    field_name="amount"
                )

        return False, None

    def validate_dataframe(
        self, df: pd.DataFrame, mapping: ColumnMapping
    ) -> ValidationResult:
        """
        Validate required columns and row-level data quality.
        """
        critical_errors: List[str] = []

        # Only check columns exist - O(1)
        if not mapping.bank_account_number or mapping.bank_account_number not in df.columns:
            critical_errors.append(CRITICAL_ERRORS["NO_ACCOUNT_COLUMN"])
            return ValidationResult(
                is_valid=False,
                critical_errors=critical_errors,
                warnings=[],
                flagged_rows=[],
                quality_report={}
            )

        if not mapping.amount or mapping.amount not in df.columns:
            critical_errors.append(CRITICAL_ERRORS["NO_AMOUNT_COLUMN"])
            return ValidationResult(
                is_valid=False,
                critical_errors=critical_errors,
                warnings=[],
                flagged_rows=[],
                quality_report={}
            )

        warnings: List[str] = []
        flagged_rows_set: Set[int] = set()

        account_col = mapping.bank_account_number
        amount_col = mapping.amount

        for position, (_, row) in enumerate(df.iterrows(), start=1):
            row_flagged = False
            has_missing, error = self._check_critical_missing_data(row, mapping, position)
            if has_missing and error:
                warnings.append(error.message)
                row_flagged = True

            account_value = row.get(account_col)
            if not _is_missing(account_value) and not self.validate_account_number(account_value):
                warnings.append(WARNING_ERRORS["INVALID_ACCOUNT"].format(row=position))
                row_flagged = True

            amount_value = row.get(amount_col)
            if not _is_missing(amount_value) and not self.validate_amount(amount_value):
                warnings.append(WARNING_ERRORS["INVALID_AMOUNT"].format(row=position))
                row_flagged = True

            if mapping.ifsc_code and mapping.ifsc_code in df.columns:
                ifsc_value = row.get(mapping.ifsc_code)
                if _is_missing(ifsc_value):
                    warnings.append(WARNING_ERRORS["MISSING_IFSC"].format(row=position))
                    row_flagged = True
                elif not self.validate_ifsc_code(ifsc_value):
                    warnings.append(WARNING_ERRORS["INVALID_IFSC_FORMAT"].format(row=position))
                    row_flagged = True

            if mapping.address and mapping.address in df.columns:
                address_value = row.get(mapping.address)
                if _is_missing(address_value):
                    warnings.append(WARNING_ERRORS["MISSING_ADDRESS"].format(row=position))
                    row_flagged = True

            if row_flagged:
                flagged_rows_set.add(position)

        duplicates = self.check_duplicate_acknowledgements(df, mapping.acknowledgement_number)
        for ack_no in duplicates:
            warnings.append(WARNING_ERRORS["DUPLICATE_ACK"].format(ack_no=ack_no))

        quality_report = self.generate_quality_report(df, mapping)

        return ValidationResult(
            is_valid=True,
            critical_errors=[],
            warnings=warnings,
            flagged_rows=sorted(flagged_rows_set),
            quality_report=quality_report
        )

    def generate_quality_report(
        self, df: pd.DataFrame, mapping: ColumnMapping
    ) -> Dict[str, Any]:
        """
        Generate counts and validity rates for mapped columns.
        """
        total_rows = len(df)

        def missing_mask(col_name: Optional[str]) -> pd.Series:
            if not col_name or col_name not in df.columns:
                return pd.Series([False] * total_rows, index=df.index)
            return df[col_name].map(_is_missing)

        def valid_count(col_name: Optional[str], validator) -> int:
            if not col_name or col_name not in df.columns:
                return 0
            return int(df[col_name].map(lambda value: (not _is_missing(value)) and validator(value)).sum())

        def rate(count: int) -> float:
            if total_rows == 0:
                return 100.0
            return round((count / total_rows) * 100, 2)

        missing_accounts = int(missing_mask(mapping.bank_account_number).sum())
        missing_amounts = int(missing_mask(mapping.amount).sum())
        missing_ifsc = int(missing_mask(mapping.ifsc_code).sum()) if mapping.ifsc_code in df.columns else 0
        missing_addresses = int(missing_mask(mapping.address).sum()) if mapping.address in df.columns else 0

        valid_accounts = valid_count(mapping.bank_account_number, self.validate_account_number)
        valid_amounts = valid_count(mapping.amount, self.validate_amount)
        valid_ifsc = valid_count(mapping.ifsc_code, self.validate_ifsc_code) if mapping.ifsc_code in df.columns else 0

        mapped_columns = [
            col for col in [
                mapping.bank_account_number,
                mapping.amount,
                mapping.ifsc_code,
                mapping.address,
            ]
            if col and col in df.columns
        ]
        if total_rows == 0 or not mapped_columns:
            completeness_rate = 100.0
        else:
            total_cells = total_rows * len(mapped_columns)
            missing_cells = int(sum(missing_mask(col).sum() for col in mapped_columns))
            completeness_rate = round(((total_cells - missing_cells) / total_cells) * 100, 2)

        duplicate_acknowledgements = len(
            self.check_duplicate_acknowledgements(df, mapping.acknowledgement_number)
        )

        return {
            "total_rows": total_rows,
            "valid_account_numbers": valid_accounts,
            "valid_ifsc_codes": valid_ifsc,
            "valid_amounts": valid_amounts,
            "missing_account_numbers": missing_accounts,
            "missing_ifsc_codes": missing_ifsc,
            "missing_addresses": missing_addresses,
            "missing_amounts": missing_amounts,
            "duplicate_acknowledgements": duplicate_acknowledgements,
            "account_number_validity_rate": rate(valid_accounts),
            "ifsc_validity_rate": rate(valid_ifsc),
            "amount_validity_rate": rate(valid_amounts),
            "data_completeness_rate": completeness_rate
        }

    def classify_error(self, error_code: str) -> ErrorCategory:
        """
        Classify an error as Critical or Warning.

        Args:
            error_code: Error code to classify

        Returns:
            ErrorCategory.CRITICAL or ErrorCategory.WARNING
        """
        if error_code in CRITICAL_ERRORS:
            return ErrorCategory.CRITICAL
        elif error_code in WARNING_ERRORS:
            return ErrorCategory.WARNING
        else:
            # Default to warning for unknown errors
            return ErrorCategory.WARNING

    def create_error_response(
        self,
        error_code: str,
        row_number: Optional[int] = None,
        field_name: Optional[str] = None,
        original_value: Optional[str] = None,
        **format_kwargs
    ) -> ErrorResponse:
        """
        Create a structured error response.

        Args:
            error_code: Error code from CRITICAL_ERRORS or WARNING_ERRORS
            row_number: Optional row number where error occurred
            field_name: Optional field name related to error
            original_value: Optional original value that caused error
            **format_kwargs: Additional format arguments for message

        Returns:
            ErrorResponse with category, code, and formatted message
        """
        category = self.classify_error(error_code)

        if error_code in CRITICAL_ERRORS:
            message = CRITICAL_ERRORS[error_code]
        elif error_code in WARNING_ERRORS:
            message = WARNING_ERRORS[error_code].format(**format_kwargs)
        else:
            message = f"Unknown error: {error_code}"

        return ErrorResponse(
            category=category,
            code=error_code,
            message=message,
            row_number=row_number,
            field_name=field_name,
            original_value=original_value
        )
