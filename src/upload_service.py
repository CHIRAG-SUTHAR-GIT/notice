"""
Upload Service for the Fraud Analysis Application.

Handles file upload validation, reading Excel/CSV files into DataFrames,
and providing preview functionality.
"""

import os
from dataclasses import dataclass
from io import BytesIO
from typing import BinaryIO, List, Optional, Union

import pandas as pd

from src.models import ValidationResult


@dataclass
class FileValidationResult:
    """Result of file validation."""
    is_valid: bool
    error_message: Optional[str] = None
    file_size_mb: float = 0.0
    file_extension: str = ""


class UploadService:
    """
    Service for handling file uploads in the fraud analysis application.
    
    Validates file types and sizes, reads Excel/CSV files into DataFrames,
    and provides preview functionality.
    """
    
    MAX_FILE_SIZE_MB: float = 50.0
    ALLOWED_EXTENSIONS: List[str] = ['.xlsx', '.xls', '.csv']
    
    def validate_file(
        self,
        file: Union[BinaryIO, BytesIO],
        filename: str
    ) -> FileValidationResult:
        """
        Validate file type and size.
        
        Args:
            file: File-like object containing the uploaded file data
            filename: Original filename with extension
            
        Returns:
            FileValidationResult with validation status and any error messages
        """
        # Get file extension
        _, ext = os.path.splitext(filename.lower())
        
        # Validate file extension
        if ext not in self.ALLOWED_EXTENSIONS:
            return FileValidationResult(
                is_valid=False,
                error_message=f"Invalid file type '{ext}'. Allowed types: {', '.join(self.ALLOWED_EXTENSIONS)}",
                file_extension=ext
            )
        
        # Get file size
        file.seek(0, 2)  # Seek to end
        file_size_bytes = file.tell()
        file.seek(0)  # Reset to beginning
        
        file_size_mb = file_size_bytes / (1024 * 1024)
        
        # Validate file size
        if file_size_mb > self.MAX_FILE_SIZE_MB:
            return FileValidationResult(
                is_valid=False,
                error_message=f"File size ({file_size_mb:.2f}MB) exceeds maximum allowed size ({self.MAX_FILE_SIZE_MB}MB)",
                file_size_mb=file_size_mb,
                file_extension=ext
            )
        
        return FileValidationResult(
            is_valid=True,
            file_size_mb=file_size_mb,
            file_extension=ext
        )

    def read_file(
        self,
        file: Union[BinaryIO, BytesIO],
        filename: str
    ) -> pd.DataFrame:
        """
        Read Excel or CSV file into a pandas DataFrame - OPTIMIZED for large files.
        
        Args:
            file: File-like object containing the file data
            filename: Original filename with extension
            
        Returns:
            DataFrame containing the file data
            
        Raises:
            ValueError: If file format is not supported or file is corrupted
        """
        _, ext = os.path.splitext(filename.lower())
        
        # Reset file position to beginning
        file.seek(0)
        
        try:
            if ext == '.csv':
                # Optimized CSV reading for large files
                # Use C engine for speed, low_memory=False for consistent dtypes
                df = pd.read_csv(
                    file, 
                    low_memory=False,
                    engine='c',  # Faster C parser
                    on_bad_lines='skip',  # Skip malformed rows
                    dtype_backend='numpy_nullable'  # Better memory handling
                )
            elif ext == '.xlsx':
                # openpyxl is required for xlsx
                df = pd.read_excel(file, engine='openpyxl')
            elif ext == '.xls':
                df = pd.read_excel(file, engine='xlrd')
            else:
                raise ValueError(f"Unsupported file format: {ext}")
            
            return df
            
        except pd.errors.EmptyDataError:
            raise ValueError("File is empty or contains no data")
        except pd.errors.ParserError as e:
            raise ValueError(f"File is corrupted or has invalid format: {str(e)}")
        except Exception as e:
            raise ValueError(f"Error reading file: {str(e)}")
    
    def get_preview(
        self,
        df: pd.DataFrame,
        rows: int = 10
    ) -> pd.DataFrame:
        """
        Return first N rows of DataFrame for preview.
        
        Args:
            df: DataFrame to preview
            rows: Number of rows to return (default: 10)
            
        Returns:
            DataFrame containing the first N rows
        """
        return df.head(rows)
    
    def validate_and_read(
        self,
        file: Union[BinaryIO, BytesIO],
        filename: str
    ) -> tuple[pd.DataFrame, FileValidationResult]:
        """
        Validate file and read into DataFrame if valid.
        
        Convenience method that combines validation and reading.
        
        Args:
            file: File-like object containing the file data
            filename: Original filename with extension
            
        Returns:
            Tuple of (DataFrame, FileValidationResult)
            DataFrame will be empty if validation fails
            
        Raises:
            ValueError: If file is corrupted or unreadable
        """
        validation_result = self.validate_file(file, filename)
        
        if not validation_result.is_valid:
            return pd.DataFrame(), validation_result
        
        df = self.read_file(file, filename)
        return df, validation_result
