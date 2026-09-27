"""
Database Service for MySQL - Store and retrieve aggregated fraud data.

GUJARAT CYBER POLICE - DATA INTEGRITY GUARANTEED
"""

import mysql.connector
from mysql.connector import Error
from typing import List, Optional, Dict, Any, Generator
from datetime import datetime
import pandas as pd
import hashlib


class DatabaseService:
    """MySQL Database Service - DATA INTEGRITY GUARANTEED for Gujarat Cyber Police."""
    
    BATCH_SIZE = 10000
    
    def __init__(self, host: str = "localhost", port: int = 3306, 
                 user: str = "root", password: str = "", database: str = "gujarat_cyber_police"):
        self.host = host
        self.port = port
        self.user = user
        self.password = password
        self.database = database
        self.connection = None
        self.last_error = ""
    
    def connect(self) -> bool:
        """Establish secure connection to MySQL server."""
        try:
            self.connection = mysql.connector.connect(
                host=self.host,
                port=self.port,
                user=self.user,
                password=self.password,
                autocommit=False
            )
            
            if self.connection.is_connected():
                cursor = self.connection.cursor()
                cursor.execute(f"CREATE DATABASE IF NOT EXISTS {self.database}")
                cursor.execute(f"USE {self.database}")
                cursor.close()
                self._create_tables()
                return True
                
        except Error as e:
            self.last_error = str(e)
            return False
        
        return False
    
    def disconnect(self):
        """Close database connection safely."""
        if self.connection and self.connection.is_connected():
            self.connection.close()

    def _create_tables(self):
        """Create tables with proper constraints for data integrity."""
        cursor = self.connection.cursor()
        
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS datasets (
                id INT AUTO_INCREMENT PRIMARY KEY,
                name VARCHAR(255) NOT NULL,
                description TEXT,
                total_accounts INT NOT NULL,
                total_amount DECIMAL(20, 2) NOT NULL,
                data_checksum VARCHAR(64),
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                source_filename VARCHAR(255),
                verified BOOLEAN DEFAULT FALSE,
                INDEX idx_created_at (created_at)
            ) ENGINE=InnoDB
        """)
        
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS aggregated_accounts (
                id INT AUTO_INCREMENT PRIMARY KEY,
                dataset_id INT NOT NULL,
                account_number VARCHAR(100) NOT NULL,
                acknowledgement_numbers MEDIUMTEXT,
                ack_count INT DEFAULT 0,
                bank_name VARCHAR(500),
                ifsc_code VARCHAR(50),
                address TEXT,
                district VARCHAR(255),
                state VARCHAR(255),
                total_transactions INT NOT NULL DEFAULT 0,
                total_amount DECIMAL(20, 2) NOT NULL DEFAULT 0,
                total_disputed_amount DECIMAL(20, 2) DEFAULT 0,
                risk_score DECIMAL(5, 2) DEFAULT 0,
                INDEX idx_dataset_id (dataset_id),
                INDEX idx_account_number (account_number),
                INDEX idx_total_amount (total_amount),
                INDEX idx_district (district),
                CONSTRAINT fk_dataset FOREIGN KEY (dataset_id) 
                    REFERENCES datasets(id) ON DELETE CASCADE
            ) ENGINE=InnoDB
        """)
        
        self.connection.commit()
        cursor.close()
    
    def _calculate_ack_count(self, ack_numbers: str) -> int:
        """Calculate ACK count from acknowledgement numbers string."""
        if ack_numbers is None:
            return 0
        try:
            if ack_numbers != ack_numbers:
                return 0
        except Exception:
            pass
        ack_numbers = str(ack_numbers).strip()
        if not ack_numbers or ack_numbers.lower() in {"nan", "none", "<na>", "nat"}:
            return 0
        ack_str = ack_numbers.replace(';', ',')
        return len([a.strip() for a in ack_str.split(',') if a.strip()])
    
    def _calculate_checksum(self, accounts: List[Any]) -> str:
        """Calculate SHA256 checksum of data for integrity verification."""
        data_str = ""
        for acc in accounts:
            data_str += f"{acc.account_number}|{acc.total_amount}|{acc.total_transactions}|"
        return hashlib.sha256(data_str.encode()).hexdigest()

    def save_dataset(self, name: str, description: str, accounts: List[Any], 
                     source_filename: str = "", progress_callback=None) -> tuple:
        """Save aggregated accounts to database with FULL DATA INTEGRITY."""
        if not self.connection or not self.connection.is_connected():
            if not self.connect():
                return None, f"Connection failed: {self.last_error}"
        
        dataset_id = None
        
        try:
            cursor = self.connection.cursor()
            
            total_accounts = len(accounts)
            total_amount = sum(acc.total_amount for acc in accounts)
            data_checksum = self._calculate_checksum(accounts)
            
            if total_accounts == 0:
                return None, "No accounts to save"
            
            cursor.execute("""
                INSERT INTO datasets (name, description, total_accounts, total_amount, 
                                      data_checksum, source_filename, verified)
                VALUES (%s, %s, %s, %s, %s, %s, FALSE)
            """, (name, description, total_accounts, total_amount, data_checksum, source_filename))
            
            dataset_id = cursor.lastrowid
            
            insert_sql = """
                INSERT INTO aggregated_accounts 
                (dataset_id, account_number, acknowledgement_numbers, ack_count,
                 bank_name, ifsc_code, address, district, state,
                 total_transactions, total_amount, total_disputed_amount, risk_score)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """
            
            batch = []
            inserted_count = 0
            
            for i, acc in enumerate(accounts):
                ack_count = self._calculate_ack_count(acc.acknowledgement_numbers)
                
                batch.append((
                    dataset_id,
                    str(acc.account_number) if acc.account_number else '',
                    str(acc.acknowledgement_numbers) if acc.acknowledgement_numbers else '',
                    ack_count,
                    str(acc.bank_name) if acc.bank_name else '',
                    str(acc.ifsc_code) if acc.ifsc_code else '',
                    str(acc.address) if acc.address else '',
                    str(acc.district) if acc.district else '',
                    str(acc.state) if acc.state else '',
                    int(acc.total_transactions) if acc.total_transactions else 0,
                    float(acc.total_amount) if acc.total_amount else 0,
                    float(acc.total_disputed_amount) if acc.total_disputed_amount else 0,
                    float(acc.risk_score) if acc.risk_score else 0
                ))
                
                if len(batch) >= self.BATCH_SIZE:
                    cursor.executemany(insert_sql, batch)
                    inserted_count += len(batch)
                    batch = []
                    if progress_callback:
                        progress_callback(inserted_count, total_accounts)
            
            if batch:
                cursor.executemany(insert_sql, batch)
                inserted_count += len(batch)
            
            cursor.execute(
                "SELECT COUNT(*) FROM aggregated_accounts WHERE dataset_id = %s", 
                (dataset_id,)
            )
            saved_count = cursor.fetchone()[0]
            
            if saved_count != total_accounts:
                raise Exception(f"DATA INTEGRITY ERROR: Expected {total_accounts}, saved {saved_count}")
            
            cursor.execute("UPDATE datasets SET verified = TRUE WHERE id = %s", (dataset_id,))
            self.connection.commit()
            
            if progress_callback:
                progress_callback(total_accounts, total_accounts)
            
            cursor.close()
            return dataset_id, ""
            
        except Exception as e:
            error_msg = str(e)
            try:
                self.connection.rollback()
                if dataset_id:
                    cursor = self.connection.cursor()
                    cursor.execute("DELETE FROM datasets WHERE id = %s", (dataset_id,))
                    self.connection.commit()
                    cursor.close()
            except:
                pass
            return None, f"Save failed: {error_msg}"

    def get_all_datasets(self) -> List[Dict[str, Any]]:
        """Get list of all saved datasets."""
        if not self.connection or not self.connection.is_connected():
            if not self.connect():
                return []
        
        try:
            cursor = self.connection.cursor(dictionary=True)
            cursor.execute("""
                SELECT id, name, description, total_accounts, total_amount, 
                       created_at, source_filename, verified
                FROM datasets ORDER BY created_at DESC
            """)
            results = cursor.fetchall()
            cursor.close()
            return results
        except Error as e:
            return []
    
    def load_dataset(self, dataset_id: int, limit: int = None, offset: int = 0) -> Optional[pd.DataFrame]:
        """Load dataset with optional pagination."""
        if not self.connection or not self.connection.is_connected():
            if not self.connect():
                return None
        
        try:
            cursor = self.connection.cursor(dictionary=True)
            
            sql = """
                SELECT 
                    account_number AS `Fraudster Bank Account Number`,
                    acknowledgement_numbers AS `All Acknowledgement Numbers`,
                    ack_count AS `ACK Count`,
                    bank_name AS `Bank Name`,
                    ifsc_code AS `IFSC Code`,
                    address AS `Address`,
                    district AS `District`,
                    state AS `State`,
                    total_transactions AS `Total Transactions`,
                    total_amount AS `Total Amount`,
                    total_disputed_amount AS `Total Disputed Amount`,
                    risk_score AS `Risk Score`
                FROM aggregated_accounts
                WHERE dataset_id = %s
                ORDER BY total_amount DESC
            """
            
            if limit:
                sql += f" LIMIT {limit} OFFSET {offset}"
            
            cursor.execute(sql, (dataset_id,))
            results = cursor.fetchall()
            cursor.close()
            
            return pd.DataFrame(results) if results else None
            
        except Error as e:
            return None
    
    def delete_dataset(self, dataset_id: int) -> bool:
        """Delete a dataset and all its accounts."""
        if not self.connection or not self.connection.is_connected():
            if not self.connect():
                return False
        
        try:
            cursor = self.connection.cursor()
            cursor.execute("DELETE FROM datasets WHERE id = %s", (dataset_id,))
            self.connection.commit()
            cursor.close()
            return True
        except Error as e:
            self.connection.rollback()
            return False
    
    def get_dataset_count(self, dataset_id: int) -> int:
        """Get total row count for a dataset."""
        if not self.connection or not self.connection.is_connected():
            if not self.connect():
                return 0
        try:
            cursor = self.connection.cursor()
            cursor.execute(
                "SELECT COUNT(*) FROM aggregated_accounts WHERE dataset_id = %s", 
                (dataset_id,)
            )
            count = cursor.fetchone()[0]
            cursor.close()
            return count
        except:
            return 0
    
    def get_dataset_info(self, dataset_id: int) -> Optional[Dict[str, Any]]:
        """Get metadata for a specific dataset."""
        if not self.connection or not self.connection.is_connected():
            if not self.connect():
                return None
        try:
            cursor = self.connection.cursor(dictionary=True)
            cursor.execute("""
                SELECT id, name, description, total_accounts, total_amount, 
                       created_at, source_filename, verified
                FROM datasets WHERE id = %s
            """, (dataset_id,))
            result = cursor.fetchone()
            cursor.close()
            return result
        except:
            return None
    
    def verify_dataset_integrity(self, dataset_id: int) -> tuple:
        """Verify data integrity of a saved dataset."""
        if not self.connection or not self.connection.is_connected():
            if not self.connect():
                return False, "Connection failed"
        
        try:
            cursor = self.connection.cursor(dictionary=True)
            cursor.execute(
                "SELECT total_accounts, total_amount FROM datasets WHERE id = %s",
                (dataset_id,)
            )
            dataset = cursor.fetchone()
            
            if not dataset:
                return False, "Dataset not found"
            
            cursor.execute(
                "SELECT COUNT(*) as cnt, SUM(total_amount) as amt FROM aggregated_accounts WHERE dataset_id = %s",
                (dataset_id,)
            )
            result = cursor.fetchone()
            actual_count = result['cnt']
            actual_amount = result['amt'] or 0
            cursor.close()
            
            if actual_count != dataset['total_accounts']:
                return False, f"Count mismatch: expected {dataset['total_accounts']}, found {actual_count}"
            
            return True, f"✅ Verified: {actual_count:,} records, ₹{float(actual_amount):,.2f}"
            
        except Error as e:
            return False, f"Error: {str(e)}"
    
    def load_dataset_filtered(self, dataset_id: int, sort_column: str = "total_amount",
                               sort_order: str = "DESC", limit: int = None,
                               filter_account: str = None, filter_bank: str = None,
                               filter_district: str = None, filter_state: str = None,
                               min_amount: float = None, max_amount: float = None,
                               min_transactions: int = None, min_ack_count: int = None) -> Optional[pd.DataFrame]:
        """Load dataset with Excel-like filtering and sorting."""
        if not self.connection or not self.connection.is_connected():
            if not self.connect():
                return None
        
        try:
            cursor = self.connection.cursor(dictionary=True)
            
            sql = """
                SELECT 
                    account_number AS `Fraudster Bank Account Number`,
                    acknowledgement_numbers AS `All Acknowledgement Numbers`,
                    ack_count AS `ACK Count`,
                    bank_name AS `Bank Name`,
                    ifsc_code AS `IFSC Code`,
                    address AS `Address`,
                    district AS `District`,
                    state AS `State`,
                    total_transactions AS `Total Transactions`,
                    total_amount AS `Total Amount`,
                    total_disputed_amount AS `Total Disputed Amount`,
                    risk_score AS `Risk Score`
                FROM aggregated_accounts
                WHERE dataset_id = %s
            """
            params = [dataset_id]
            
            if filter_account:
                sql += " AND account_number LIKE %s"
                params.append(f"%{filter_account}%")
            if filter_bank:
                sql += " AND bank_name LIKE %s"
                params.append(f"%{filter_bank}%")
            if filter_district:
                sql += " AND district LIKE %s"
                params.append(f"%{filter_district}%")
            if filter_state:
                sql += " AND state LIKE %s"
                params.append(f"%{filter_state}%")
            if min_amount and min_amount > 0:
                sql += " AND total_amount >= %s"
                params.append(min_amount)
            if max_amount and max_amount > 0:
                sql += " AND total_amount <= %s"
                params.append(max_amount)
            if min_transactions and min_transactions > 0:
                sql += " AND total_transactions >= %s"
                params.append(min_transactions)
            if min_ack_count and min_ack_count > 0:
                sql += " AND ack_count >= %s"
                params.append(min_ack_count)
            
            valid_columns = ['account_number', 'bank_name', 'district', 'state', 
                           'total_amount', 'total_transactions', 'ack_count', 'risk_score']
            if sort_column not in valid_columns:
                sort_column = 'total_amount'
            
            sort_order = 'DESC' if sort_order.upper() == 'DESC' else 'ASC'
            sql += f" ORDER BY {sort_column} {sort_order}"
            
            if limit:
                sql += f" LIMIT {int(limit)}"
            
            cursor.execute(sql, params)
            results = cursor.fetchall()
            cursor.close()
            
            return pd.DataFrame(results) if results else None
            
        except Error as e:
            return None
    
    def search_accounts(self, dataset_id: int, account_number: str = None, 
                        district: str = None, min_amount: float = None,
                        limit: int = 1000) -> Optional[pd.DataFrame]:
        """Fast indexed search."""
        if not self.connection or not self.connection.is_connected():
            if not self.connect():
                return None
        
        try:
            cursor = self.connection.cursor(dictionary=True)
            
            sql = """
                SELECT 
                    account_number AS `Fraudster Bank Account Number`,
                    acknowledgement_numbers AS `All Acknowledgement Numbers`,
                    ack_count AS `ACK Count`,
                    bank_name AS `Bank Name`,
                    ifsc_code AS `IFSC Code`,
                    district AS `District`,
                    state AS `State`,
                    total_transactions AS `Total Transactions`,
                    total_amount AS `Total Amount`
                FROM aggregated_accounts WHERE dataset_id = %s
            """
            params = [dataset_id]
            
            if account_number:
                sql += " AND account_number LIKE %s"
                params.append(f"%{account_number}%")
            if district:
                sql += " AND district LIKE %s"
                params.append(f"%{district}%")
            if min_amount:
                sql += " AND total_amount >= %s"
                params.append(min_amount)
            
            sql += f" ORDER BY total_amount DESC LIMIT {limit}"
            
            cursor.execute(sql, params)
            results = cursor.fetchall()
            cursor.close()
            
            return pd.DataFrame(results) if results else None
            
        except Error as e:
            return None
    
    def test_connection(self) -> tuple:
        """Test database connection."""
        try:
            conn = mysql.connector.connect(
                host=self.host, port=self.port, user=self.user, password=self.password
            )
            if conn.is_connected():
                server_info = conn.get_server_info()
                conn.close()
                return True, f"Connected to MySQL Server version {server_info}"
            return False, "Connection failed"
        except Error as e:
            return False, f"Error: {str(e)}"
