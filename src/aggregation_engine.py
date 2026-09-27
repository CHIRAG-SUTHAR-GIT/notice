"""
Aggregation Engine - FAST but with ALL features preserved.
"""

from typing import List
import pandas as pd
import numpy as np

from src.models import AggregatedAccount, ColumnMapping


class AggregationEngine:
    """Engine for aggregating transaction data - FAST with all features."""
    
    def get_most_common(self, series: pd.Series) -> str:
        """Return the most common non-null value in a series."""
        non_null = series.dropna()
        non_empty = non_null[non_null.astype(str).str.strip() != '']
        
        if len(non_empty) == 0:
            return ""
        
        mode_result = non_empty.mode()
        if len(mode_result) == 0:
            return ""
        
        return str(mode_result.iloc[0])
    
    def calculate_risk_score(self, transaction_count: int, total_amount: float) -> float:
        """Calculate risk score based on transactions and amount."""
        transaction_weight = 0.4
        amount_weight = 0.6
        
        transaction_score = min(transaction_count / 100.0, 1.0) * 100
        amount_score = min(total_amount / 10_000_000.0, 1.0) * 100
        
        risk_score = (transaction_weight * transaction_score + 
                      amount_weight * amount_score)
        
        return round(risk_score, 2)
    
    def aggregate_by_account(
        self, 
        df: pd.DataFrame, 
        mapping: ColumnMapping
    ) -> List[AggregatedAccount]:
        """
        Group transactions by account number - FAST.
        Collects ALL unique ACK numbers for each account.
        """
        if df.empty:
            return []
        
        account_col = mapping.bank_account_number
        if not account_col or account_col not in df.columns:
            return []
        
        # Filter out null/empty account numbers
        df = df.copy()
        df[account_col] = df[account_col].astype(str).str.strip()
        df = df[~df[account_col].isin(['', 'nan', 'None', 'NaN', '<NA>'])]
        
        if df.empty:
            return []
        
        # Get column names
        amount_col = mapping.amount
        disputed_col = mapping.disputed_amount
        bank_name_col = mapping.bank_name
        ifsc_col = mapping.ifsc_code
        address_col = mapping.address
        ack_col = mapping.acknowledgement_number
        district_col = mapping.district
        state_col = mapping.state
        
        # Ensure numeric columns
        if amount_col and amount_col in df.columns:
            df[amount_col] = pd.to_numeric(df[amount_col], errors='coerce').fillna(0)
        if disputed_col and disputed_col in df.columns:
            df[disputed_col] = pd.to_numeric(df[disputed_col], errors='coerce').fillna(0)
        
        # Build aggregation dict using named aggregation (FAST - no lambdas)
        agg_dict = {'_count': (account_col, 'count')}
        
        if amount_col and amount_col in df.columns:
            agg_dict['total_amount'] = (amount_col, 'sum')
        
        if disputed_col and disputed_col in df.columns:
            agg_dict['total_disputed'] = (disputed_col, 'sum')
        
        # Use 'first' for text columns
        if bank_name_col and bank_name_col in df.columns:
            agg_dict['bank_name'] = (bank_name_col, 'first')
        if ifsc_col and ifsc_col in df.columns:
            agg_dict['ifsc_code'] = (ifsc_col, 'first')
        if address_col and address_col in df.columns:
            agg_dict['address'] = (address_col, 'first')
        if district_col and district_col in df.columns:
            agg_dict['district'] = (district_col, 'first')
        if state_col and state_col in df.columns:
            agg_dict['state'] = (state_col, 'first')
        
        # Single fast groupby for numeric aggregations
        result = df.groupby(account_col, sort=False).agg(**agg_dict).reset_index()
        
        # Collect ALL unique ACK numbers per account (separate operation for speed)
        ack_dict = {}
        if ack_col and ack_col in df.columns:
            # Group and collect unique ACK numbers
            ack_grouped = df.groupby(account_col)[ack_col].apply(
                lambda x: ';'.join(x.dropna().astype(str).unique())
            )
            ack_dict = ack_grouped.to_dict()
        
        # Build result list
        aggregated_accounts = []
        
        for i in range(len(result)):
            row = result.iloc[i]
            acc_num = str(row[account_col])
            
            total_transactions = int(row['_count'])
            total_amount = float(row.get('total_amount', 0) or 0)
            total_disputed = float(row.get('total_disputed', 0) or 0)
            
            # Get all unique ACK numbers for this account
            ack_numbers = ack_dict.get(acc_num, '')
            
            risk_score = self.calculate_risk_score(total_transactions, total_amount)
            
            aggregated_accounts.append(AggregatedAccount(
                account_number=acc_num,
                bank_name=str(row.get('bank_name', '') or ''),
                ifsc_code=str(row.get('ifsc_code', '') or ''),
                address=str(row.get('address', '') or ''),
                district=str(row.get('district', '') or ''),
                state=str(row.get('state', '') or ''),
                total_transactions=total_transactions,
                acknowledgement_numbers=ack_numbers,
                total_amount=total_amount,
                total_disputed_amount=total_disputed,
                risk_score=risk_score
            ))
        
        return aggregated_accounts
    
    def sort_results(
        self, 
        accounts: List[AggregatedAccount]
    ) -> List[AggregatedAccount]:
        """Sort by total amount (desc), then by transaction count (desc)."""
        return sorted(
            accounts,
            key=lambda x: (-x.total_amount, -x.total_transactions)
        )
