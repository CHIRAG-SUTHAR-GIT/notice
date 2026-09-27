"""
Dashboard module for the Fraud Analysis Application.

Provides statistics calculation, search, and filtering capabilities
for aggregated fraud transaction data.
"""

from typing import List, Optional
from datetime import datetime

from src.models import AggregatedAccount, ProcessingStats


class Dashboard:
    """Dashboard for viewing and filtering fraud analysis results."""
    
    def calculate_statistics(
        self, 
        accounts: List[AggregatedAccount],
        total_input_rows: int,
        input_filename: str = "",
        rows_with_errors: int = 0
    ) -> ProcessingStats:
        """
        Calculate summary statistics from aggregated accounts.
        
        Args:
            accounts: List of aggregated account data.
            total_input_rows: Total number of rows in the input file.
            input_filename: Name of the input file.
            rows_with_errors: Number of rows that had errors.
            
        Returns:
            ProcessingStats object with calculated statistics.
        """
        if not accounts:
            return ProcessingStats(
                total_input_rows=total_input_rows,
                rows_processed=total_input_rows - rows_with_errors,
                rows_with_errors=rows_with_errors,
                unique_accounts=0,
                total_fraud_amount=0.0,
                total_disputed_amount=0.0,
                average_amount_per_account=0.0,
                top_accounts_by_amount=[],
                processing_timestamp=datetime.now(),
                input_filename=input_filename
            )
        
        # Calculate totals
        total_fraud_amount = sum(acc.total_amount for acc in accounts)
        total_disputed_amount = sum(acc.total_disputed_amount for acc in accounts)
        unique_accounts = len(accounts)
        
        # Calculate average
        average_amount = total_fraud_amount / unique_accounts if unique_accounts > 0 else 0.0
        
        # Get top 20 accounts by amount
        sorted_accounts = sorted(accounts, key=lambda x: -x.total_amount)
        top_accounts = sorted_accounts[:20]
        
        return ProcessingStats(
            total_input_rows=total_input_rows,
            rows_processed=total_input_rows - rows_with_errors,
            rows_with_errors=rows_with_errors,
            unique_accounts=unique_accounts,
            total_fraud_amount=total_fraud_amount,
            total_disputed_amount=total_disputed_amount,
            average_amount_per_account=average_amount,
            top_accounts_by_amount=top_accounts,
            processing_timestamp=datetime.now(),
            input_filename=input_filename
        )
    
    def search_accounts(
        self, 
        accounts: List[AggregatedAccount], 
        query: str
    ) -> List[AggregatedAccount]:
        """
        Search for accounts by account number.
        
        Args:
            accounts: List of aggregated accounts to search.
            query: Search query string to match against account numbers.
            
        Returns:
            List of accounts where account number contains the query.
        """
        if not query or not query.strip():
            return accounts
        
        query = query.strip()
        
        return [
            acc for acc in accounts 
            if query in acc.account_number
        ]
    
    def filter_by_min_transactions(
        self, 
        accounts: List[AggregatedAccount], 
        min_transactions: int
    ) -> List[AggregatedAccount]:
        """
        Filter accounts by minimum transaction count.
        
        Args:
            accounts: List of aggregated accounts to filter.
            min_transactions: Minimum number of transactions required.
            
        Returns:
            List of accounts with transaction count >= min_transactions.
        """
        if min_transactions <= 0:
            return accounts
        
        return [
            acc for acc in accounts 
            if acc.total_transactions >= min_transactions
        ]
    
    def filter_by_min_amount(
        self, 
        accounts: List[AggregatedAccount], 
        min_amount: float
    ) -> List[AggregatedAccount]:
        """
        Filter accounts by minimum total amount.
        
        Args:
            accounts: List of aggregated accounts to filter.
            min_amount: Minimum total amount required.
            
        Returns:
            List of accounts with total_amount >= min_amount.
        """
        if min_amount <= 0:
            return accounts
        
        return [
            acc for acc in accounts 
            if acc.total_amount >= min_amount
        ]
    
    def get_flagged_rows(
        self, 
        accounts: List[AggregatedAccount],
        flagged_account_numbers: List[str]
    ) -> List[AggregatedAccount]:
        """
        Get accounts that were flagged during processing.
        
        Args:
            accounts: List of all aggregated accounts.
            flagged_account_numbers: List of account numbers that were flagged.
            
        Returns:
            List of accounts that match the flagged account numbers.
        """
        flagged_set = set(flagged_account_numbers)
        return [
            acc for acc in accounts 
            if acc.account_number in flagged_set
        ]
