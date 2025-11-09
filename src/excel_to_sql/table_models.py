"""
Table Models Module
Contains data models for table detection and processing.
"""

from dataclasses import dataclass
from typing import Dict, Optional


@dataclass
class TableLocation:
    """Represents the location and bounds of a detected table"""
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    width_row: int  # Row containing width headers
    height_col: int  # Column containing height headers
    table_type: str = "standard"  # "standard" or "other"
    price_cols: Dict = None  # For other type tables
    
    def __post_init__(self):
        if self.price_cols is None:
            self.price_cols = {}
