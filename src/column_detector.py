"""
Column Detector for the Fraud Analysis Application.

Uses fuzzy matching to detect and map column headers from various naming conventions
to standardized column types.
"""

import re
from typing import Dict, List, Optional, Tuple

from rapidfuzz import fuzz

from src.models import ColumnMapping


class ColumnDetector:
    """
    Detects and maps column headers using fuzzy matching.
    
    Supports various naming conventions for fraud transaction data columns
    and provides confidence scores for automatic suggestions.
    """
    
    SIMILARITY_THRESHOLD: float = 0.80
    
    # Known column variants for each standardized column type
    # Requirements: 2.6-2.13
    COLUMN_VARIANTS: Dict[str, List[str]] = {
        'serial_number': [
            'sr no', 'sr.no', 'serial no', 's.no', 'sno', 
            'serial number', '#'
        ],
        'acknowledgement_number': [
            'acknowledgement no', 'ack no', 'ackno', 'ack',
            'acknowledgment no', 'acknowledgement number',
            'acknowledgment number', 'ref no', 'reference no'
        ],
        'bank_account_number': [
            'bank account no', 'bank ac no', 'bank a/c no', 'ac no',
            'a/c no', 'account no', 'account number',
            'bank account number', 'beneficiary account', 'beneficiary ac'
        ],
        'ifsc_code': [
            'ifsc code', 'ifsc', 'bank code'
        ],
        'address': [
            'address', 'beneficiary address', 'account holder address', 'location'
        ],
        'amount': [
            'amount', 'transaction amount', 'txn amount', 
            'transfer amount', 'fraud amount'
        ],
        'disputed_amount': [
            'disputed amount', 'disputed', 'claim amount',
            'disputed amt', 'chargeback amount'
        ],
        'bank_name': [
            'bank name', 'bank', 'beneficiary bank', 'receiving bank',
            'bank fis', 'bank fi', 'bankfis'
        ],
        'district': [
            'district', 'dist', 'district name'
        ],
        'state': [
            'state', 'state name', 'province'
        ]
    }
    
    def normalize_header(self, header: str) -> str:
        """
        Normalize a header string for comparison.
        
        Strips whitespace, converts to lowercase, and removes special characters.
        
        Args:
            header: The raw header string to normalize.
            
        Returns:
            Normalized header string.
            
        Requirements: 2.2
        """
        if not isinstance(header, str):
            header = str(header)
        
        # Strip leading/trailing whitespace
        normalized = header.strip()
        
        # Convert to lowercase
        normalized = normalized.lower()
        
        # Remove special characters except spaces (keep alphanumeric and spaces)
        # Replace common separators with spaces first
        normalized = normalized.replace('_', ' ')
        normalized = normalized.replace('-', ' ')
        normalized = normalized.replace('.', ' ')
        normalized = normalized.replace('/', ' ')
        
        # Remove any remaining special characters (keep letters, numbers, spaces)
        normalized = re.sub(r'[^a-z0-9\s]', '', normalized)
        
        # Collapse multiple spaces into single space
        normalized = re.sub(r'\s+', ' ', normalized)
        
        # Final strip
        normalized = normalized.strip()
        
        return normalized
    
    def calculate_similarity(self, header: str, variant: str) -> float:
        """
        Calculate fuzzy similarity score between a header and a variant.
        
        Uses rapidfuzz for efficient fuzzy string matching.
        
        Args:
            header: The normalized header string.
            variant: The known variant to compare against.
            
        Returns:
            Similarity score between 0.0 and 1.0.
            
        Requirements: 2.1
        """
        # Use token_sort_ratio for better matching of reordered words
        # and partial_ratio for substring matching
        ratio = fuzz.ratio(header, variant) / 100.0
        token_sort = fuzz.token_sort_ratio(header, variant) / 100.0
        partial = fuzz.partial_ratio(header, variant) / 100.0
        
        # Return the maximum of the three methods for best match
        return max(ratio, token_sort, partial)
    
    def _find_best_match(
        self, 
        normalized_header: str,
        original_header: str = ""
    ) -> Tuple[Optional[str], float, List[Tuple[str, float]]]:
        """
        Find the best matching column type for a normalized header.
        
        Args:
            normalized_header: The normalized header string.
            original_header: The original header string (for exact matching).
            
        Returns:
            Tuple of (best_column_type, confidence_score, all_matches_above_threshold)
        """
        best_match: Optional[str] = None
        best_score: float = 0.0
        all_matches: List[Tuple[str, float]] = []
        
        # Also normalize the original header for comparison (lowercase, strip)
        original_lower = original_header.lower().strip() if original_header else ""
        
        # First pass: look for exact matches (highest priority)
        for column_type, variants in self.COLUMN_VARIANTS.items():
            for variant in variants:
                # Check for exact match with original header (case-insensitive)
                if original_lower == variant.lower():
                    # Exact match - 100% confidence
                    all_matches.append((column_type, 1.0))
                    if 1.0 > best_score:
                        best_score = 1.0
                        best_match = column_type
        
        # If we found an exact match, return it immediately
        if best_score == 1.0:
            return best_match, best_score, all_matches
        
        # Second pass: fuzzy matching (only if no exact match found)
        for column_type, variants in self.COLUMN_VARIANTS.items():
            for variant in variants:
                # Use ratio (not partial_ratio) to avoid substring matching issues
                ratio = fuzz.ratio(normalized_header, variant) / 100.0
                token_sort = fuzz.token_sort_ratio(normalized_header, variant) / 100.0
                
                # Take the max of ratio and token_sort, but NOT partial_ratio
                # to avoid 'disputed amount' matching 'amount'
                score = max(ratio, token_sort)
                
                if score >= self.SIMILARITY_THRESHOLD:
                    all_matches.append((column_type, score))
                    
                    if score > best_score:
                        best_score = score
                        best_match = column_type
        
        return best_match, best_score, all_matches
    
    def detect_columns(self, headers: List[str]) -> ColumnMapping:
        """
        Detect and map columns using fuzzy matching.
        
        Maps input headers to standardized column types based on similarity
        to known variants. Flags ambiguous mappings for user confirmation.
        
        Args:
            headers: List of column header strings from the input file.
            
        Returns:
            ColumnMapping with detected columns, confidence scores,
            and any ambiguous mappings.
            
        Requirements: 2.1, 2.2, 2.3, 2.4
        """
        mapping = ColumnMapping()
        confidence_scores: Dict[str, float] = {}
        ambiguous_mappings: Dict[str, List[str]] = {}
        
        # Track which column types have been assigned
        assigned_types: Dict[str, str] = {}  # column_type -> original_header
        
        for header in headers:
            normalized = self.normalize_header(header)
            best_match, best_score, all_matches = self._find_best_match(normalized, header)
            
            if best_match is None:
                continue
            
            # Check for ambiguous mappings (multiple column types match)
            unique_types = list(set(match[0] for match in all_matches))
            if len(unique_types) > 1:
                # Multiple potential matches - flag for user confirmation
                ambiguous_mappings[header] = unique_types
            
            # Check if this column type is already assigned
            if best_match in assigned_types:
                # Compare scores and keep the better match
                existing_header = assigned_types[best_match]
                existing_score = confidence_scores.get(best_match, 0.0)
                
                if best_score > existing_score:
                    # New header is a better match
                    assigned_types[best_match] = header
                    confidence_scores[best_match] = best_score
                    setattr(mapping, best_match, header)
            else:
                # First match for this column type
                assigned_types[best_match] = header
                confidence_scores[best_match] = best_score
                setattr(mapping, best_match, header)
        
        mapping.confidence_scores = confidence_scores
        mapping.ambiguous_mappings = ambiguous_mappings
        
        return mapping
    
    def get_unmapped_headers(
        self, 
        headers: List[str], 
        mapping: ColumnMapping
    ) -> List[str]:
        """
        Get headers that were not mapped to any column type.
        
        Args:
            headers: Original list of headers.
            mapping: The column mapping result.
            
        Returns:
            List of headers that were not mapped.
        """
        mapped_headers = set()
        
        for field in [
            'serial_number', 'acknowledgement_number', 'bank_account_number',
            'ifsc_code', 'address', 'amount', 'disputed_amount', 'bank_name',
            'district', 'state'
        ]:
            value = getattr(mapping, field)
            if value is not None:
                mapped_headers.add(value)
        
        return [h for h in headers if h not in mapped_headers]
    
    def validate_required_columns(self, mapping: ColumnMapping) -> List[str]:
        """
        Validate that required columns are present in the mapping.
        
        Args:
            mapping: The column mapping to validate.
            
        Returns:
            List of missing required column names.
        """
        missing = []
        
        if mapping.bank_account_number is None:
            missing.append('bank_account_number')
        
        if mapping.amount is None:
            missing.append('amount')
        
        return missing
