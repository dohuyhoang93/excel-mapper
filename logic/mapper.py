"""
Mapping logic for suggesting column relationships between source and destination Excel files
"""
import difflib
from typing import List, Dict, Optional, Tuple
# from config import Config

class ColumnMapper:
    """Handles intelligent column mapping suggestions"""
    
    def __init__(self):
        self.common_mappings = {} # Config.COMMON_MAPPINGS
    
    def suggest_mapping(self, source_column: str, destination_columns: List[str]) -> Optional[str]:
        """
        Suggest the best matching destination column for a source column
        
        Args:
            source_column: Name of the source column
            destination_columns: List of available destination columns
            
        Returns:
            Best matching destination column name or None
        """
        if not source_column or not destination_columns:
            return None
        
        source_lower = source_column.lower().strip()
        
        # 1. Exact match (case insensitive)
        exact_match = self._find_exact_match(source_lower, destination_columns)
        if exact_match:
            return exact_match
        
        # 2. Common keyword mapping
        keyword_match = self._find_keyword_match(source_lower, destination_columns)
        if keyword_match:
            return keyword_match
        
        # 3. Partial string matching
        partial_match = self._find_partial_match(source_lower, destination_columns)
        if partial_match:
            return partial_match
        
        # 4. Fuzzy matching using difflib
        fuzzy_match = self._find_fuzzy_match(source_column, destination_columns)
        if fuzzy_match:
            return fuzzy_match
        
        return None
    
    def _find_exact_match(self, source_lower: str, destination_columns: List[str]) -> Optional[str]:
        """Find exact case-insensitive match"""
        for dest_col in destination_columns:
            if dest_col.lower().strip() == source_lower:
                return dest_col
        return None
    
    def _find_keyword_match(self, source_lower: str, destination_columns: List[str]) -> Optional[str]:
        """Find match based on common keywords"""
        for keyword, possible_dest in self.common_mappings.items():
            if keyword in source_lower:
                for dest_col in destination_columns:
                    if any(possible.lower() in dest_col.lower() for possible in possible_dest):
                        return dest_col
        return None
    
    def _find_partial_match(self, source_lower: str, destination_columns: List[str]) -> Optional[str]:
        """Find partial string match"""
        best_match = None
        best_score = 0
        
        for dest_col in destination_columns:
            dest_lower = dest_col.lower().strip()
            
            # Check if source is contained in destination or vice versa
            if source_lower in dest_lower:
                score = len(source_lower) / len(dest_lower)
                if score > best_score:
                    best_score = score
                    best_match = dest_col
            elif dest_lower in source_lower:
                score = len(dest_lower) / len(source_lower)
                if score > best_score:
                    best_score = score
                    best_match = dest_col
        
        # Only return if score is above threshold
        return best_match if best_score > 0.3 else None
    
    def _find_fuzzy_match(self, source_column: str, destination_columns: List[str], threshold: float = 0.6) -> Optional[str]:
        """Find fuzzy match using sequence matching"""
        best_matches = difflib.get_close_matches(
            source_column, 
            destination_columns, 
            n=1, 
            cutoff=threshold
        )
        return best_matches[0] if best_matches else None
    
    def validate_mappings(self, mappings: Dict[str, str]) -> Tuple[bool, List[str]]:
        """
        Validate column mappings for duplicates and conflicts
        
        Args:
            mappings: Dictionary of source -> destination column mappings
            
        Returns:
            Tuple of (is_valid, list_of_errors)
        """
        errors = []
        
        # Check for duplicate destination columns
        destination_columns = list(mappings.values())
        duplicates = []
        
        for dest_col in set(destination_columns):
            if destination_columns.count(dest_col) > 1:
                duplicates.append(dest_col)
        
        if duplicates:
            errors.append(f"Duplicate destination columns: {', '.join(duplicates)}")
        
        # Check for empty mappings
        empty_mappings = [src for src, dest in mappings.items() if not dest.strip()]
        if empty_mappings:
            errors.append(f"Empty destination mappings for: {', '.join(empty_mappings)}")
        
        return len(errors) == 0, errors
    
    def get_mapping_suggestions(self, source_columns: List[str], destination_columns: List[str]) -> Dict[str, str]:
        """
        Get mapping suggestions for all source columns
        
        Args:
            source_columns: List of source column names
            destination_columns: List of destination column names
            
        Returns:
            Dictionary of suggested mappings
        """
        suggestions = {}
        used_destinations = set()
        
        # First pass: exact and keyword matches
        for source_col in source_columns:
            suggestion = self.suggest_mapping(source_col, destination_columns)
            if suggestion and suggestion not in used_destinations:
                suggestions[source_col] = suggestion
                used_destinations.add(suggestion)
        
        # Second pass: fuzzy matches for unmapped columns
        remaining_destinations = [col for col in destination_columns if col not in used_destinations]
        unmapped_sources = [col for col in source_columns if col not in suggestions]
        
        for source_col in unmapped_sources:
            suggestion = self._find_fuzzy_match(source_col, remaining_destinations, threshold=0.5)
            if suggestion:
                suggestions[source_col] = suggestion
                remaining_destinations.remove(suggestion)
        
        return suggestions
    
    def add_custom_mapping(self, keyword: str, destination_options: List[str]):
        """Add custom keyword mapping"""
        self.common_mappings[keyword.lower()] = destination_options
    
    def get_mapping_confidence(self, source_column: str, destination_column: str) -> float:
        """
        Calculate confidence score for a mapping (0.0 to 1.0)
        
        Args:
            source_column: Source column name
            destination_column: Destination column name
            
        Returns:
            Confidence score between 0.0 and 1.0
        """
        if not source_column or not destination_column:
            return 0.0
        
        source_lower = source_column.lower().strip()
        dest_lower = destination_column.lower().strip()
        
        # Exact match
        if source_lower == dest_lower:
            return 1.0
        
        # Keyword match
        for keyword, possible_dest in self.common_mappings.items():
            if keyword in source_lower:
                if any(possible.lower() in dest_lower for possible in possible_dest):
                    return 0.9
        
        # Partial match
        if source_lower in dest_lower or dest_lower in source_lower:
            overlap = max(len(source_lower), len(dest_lower))
            total = len(source_lower) + len(dest_lower)
            return 0.7 * (overlap / total)
        
        # Fuzzy match
        similarity = difflib.SequenceMatcher(None, source_lower, dest_lower).ratio()
        return similarity * 0.6  # Max 0.6 for fuzzy matches