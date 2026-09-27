"""
Data Processor - FAST but with ALL features preserved.
"""

import pandas as pd
import numpy as np
from typing import Optional
from src.models import ColumnMapping


class DataProcessor:
    """
    Processes and cleans transaction data - OPTIMIZED.
    All features preserved, just faster implementation.
    """
    
    CURRENCY_SYMBOLS = ['₹', '$', '£', '€', 'Rs.', 'Rs', 'INR', 'USD']
    
    def _normalize_empty_values(self, df: pd.DataFrame) -> pd.DataFrame:
        """Trim text cells and convert blank/null-like strings to missing values."""
        df = df.copy()
        str_cols = df.select_dtypes(include=['object', 'string']).columns
        for col in str_cols:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace(['nan', 'NaN', 'None', '<NA>', 'NaT', ''], pd.NA)
        return df

    def clean_dataframe(self, df: pd.DataFrame, mapping: ColumnMapping) -> pd.DataFrame:
        """
        Apply all cleaning operations - FAST but complete.
        """
        df = df.copy()
        
        # Step 1: Trim text and remove rows that are empty after trimming.
        df = self._normalize_empty_values(df)
        df = self.remove_empty_rows(df)
        
        # Step 3: Standardize account numbers
        if mapping.bank_account_number and mapping.bank_account_number in df.columns:
            df[mapping.bank_account_number] = (
                df[mapping.bank_account_number]
                .astype(str)
                .str.replace(r'[\s\-]', '', regex=True)
                .replace(['nan', 'None', ''], pd.NA)
            )
        
        # Step 4: Parse amounts
        if mapping.amount and mapping.amount in df.columns:
            df[mapping.amount] = self._parse_amounts_fast(df[mapping.amount])
        
        if mapping.disputed_amount and mapping.disputed_amount in df.columns:
            df[mapping.disputed_amount] = self._parse_amounts_fast(df[mapping.disputed_amount])
        
        return df
    
    def _parse_amounts_fast(self, series: pd.Series) -> pd.Series:
        """Parse amounts - vectorized and fast."""
        result = series.astype(str).str.strip()
        missing_mask = series.isna() | result.str.lower().isin(['', 'nan', 'none', '<na>', 'nat'])
        
        # Remove currency symbols
        for symbol in self.CURRENCY_SYMBOLS:
            result = result.str.replace(symbol, '', regex=False)
        
        # Remove commas
        result = result.str.replace(',', '', regex=False)
        result = result.str.strip()
        
        # Convert invalid non-empty values to 0 while preserving true missing values.
        parsed = pd.to_numeric(result, errors='coerce')
        return parsed.mask(~missing_mask & parsed.isna(), 0.0)
    
    # Keep these for compatibility
    def remove_empty_rows(self, df: pd.DataFrame) -> pd.DataFrame:
        normalized = self._normalize_empty_values(df)
        return normalized.dropna(how='all').reset_index(drop=True)
    
    def trim_whitespace(self, df: pd.DataFrame) -> pd.DataFrame:
        return self._normalize_empty_values(df)
    
    def standardize_account_numbers_vectorized(self, series: pd.Series) -> pd.Series:
        result = series.astype(str)
        result = result.str.replace(r'[\s\-]', '', regex=True)
        result = result.replace(['nan', 'None', ''], pd.NA)
        return result
    
    def standardize_account_number(self, account: Optional[str]) -> str:
        if account is None or account == 'nan' or account == 'None':
            return ''
        import re
        return re.sub(r'[\s\-]', '', str(account))
    
    def parse_amounts_vectorized(self, series: pd.Series) -> pd.Series:
        return self._parse_amounts_fast(series)
    
    def parse_amount(self, amount_str: Optional[str]) -> float:
        if amount_str is None or amount_str == 'nan' or amount_str == 'None' or amount_str == '':
            return 0.0
        
        amount_str = str(amount_str).strip()
        if not amount_str:
            return 0.0
        
        cleaned = amount_str
        for symbol in self.CURRENCY_SYMBOLS:
            cleaned = cleaned.replace(symbol, '')
        
        cleaned = cleaned.replace(',', '').strip()
        
        try:
            return float(cleaned)
        except (ValueError, TypeError):
            return 0.0
