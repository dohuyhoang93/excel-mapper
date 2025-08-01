"""
Handles intelligent column mapping suggestions between source and destination columns.
"""
import re
from typing import List, Set

class ColumnMapper:
    """Provides methods to suggest column mappings based on name similarity."""

    def _normalize_and_tokenize(self, text: str) -> Set[str]:
        """
        Normalizes a column header string by converting it to lowercase,
        removing special characters, and splitting it into a set of words (tokens).
        """
        # Replace various whitespace characters and separators with a single space
        text = re.sub(r'[\s\u3000_\-]+', ' ', text)
        # Remove common bracketing characters
        text = re.sub(r'[()\[\]{}]', '', text)
        return set(text.lower().strip().split())

    def suggest_mapping(self, source_col: str, dest_cols: List[str]) -> str:
        """
        Suggests the best matching destination column for a given source column
        using a scoring-based algorithm.

        Args:
            source_col: The name of the source column.
            dest_cols: A list of available destination column names.

        Returns:
            The name of the best matching destination column, or an empty string if no
            suitable match is found.
        """
        # Do not suggest mappings for unnamed or generic source columns
        if str(source_col).startswith('Column_'):
            return ""

        source_tokens = self._normalize_and_tokenize(source_col)
        if not source_tokens:
            return ""

        best_match = ""
        max_score = 0

        # A map of common keywords to boost scores for semantic matches
        keywords_map = {
            'content': 'contents', 'purpose': 'purpose', 'amount': 'amount',
            'vat': 'vat', 'currency': 'currency', 'date': 'trading date',
            'no': 'no.', 'number': 'no.', 'code': 'code reference', 'total': 'sub total'
        }

        for dest_col in dest_cols:
            current_score = 0
            dest_tokens = self._normalize_and_tokenize(dest_col)
            if not dest_tokens:
                continue

            # Perfect match gives a very high score
            if source_tokens == dest_tokens:
                current_score = 100

            # Score based on the number of common words
            common_tokens = source_tokens.intersection(dest_tokens)
            current_score += len(common_tokens) * 50

            # Boost score for known keyword synonyms
            for key, value in keywords_map.items():
                if key in source_tokens and value in dest_tokens:
                    current_score += 40

            # Boost score if one name is a substring of the other (after normalization)
            source_norm_str = "".join(sorted(list(source_tokens)))
            dest_norm_str = "".join(sorted(list(dest_tokens)))
            if source_norm_str in dest_norm_str or dest_norm_str in source_norm_str:
                current_score += 20

            if current_score > max_score:
                max_score = current_score
                best_match = dest_col

        return best_match
