"""
Other Table Handler Module
Handles other-type table detection and price extraction for the quotation system.
"""

import re
from typing import Optional, Tuple
from table_models import TableLocation


class OtherTableHandler:
    """Handles other-type table detection and price extraction"""
    
    def __init__(self, is_inch_value_func=None):
        self.is_inch_value = is_inch_value_func
    
    def get_other_table_bounderies(self, sheet, start_row: int, start_col: int) -> Optional[TableLocation]:
        """Get boundaries of other-type table structure using keyword recognition"""
        
        # Look for width row (containing inch values like "4"", "6"")
        width_row = None
        for row in range(start_row, start_row + 5):
            cell_value = sheet.cell(row, start_col).value
            if self.is_inch_value(cell_value):
                width_row = row
                break
        
        if width_row is None:
            return None
        
        # Find height column (containing any values)
        height_col = None
        for col in range(start_col, start_col + 5):
            cell_value = sheet.cell(start_row, col).value
            if cell_value is not None and str(cell_value).strip():
                height_col = col
                break
        
        if height_col is None:
            return None
        
        # Find the extent of width and height columns
        # Find the extent of height rows
        end_row = None
        for row in range(width_row + 2, sheet.max_row + 1, 2):
            cell_value = sheet.cell(row, start_col).value
            if not self.is_inch_value(cell_value):
                end_row = row - 1
                break
            end_row = row + 2

        # Find the end of height column
        end_col = None
        for col in range(height_col, sheet.max_column + 1):
            cell_value = sheet.cell(start_row, col).value
            if cell_value is None or not str(cell_value).strip():
                end_col = col - 1
                break
            end_col = col + 1
        
        # Validate table boundaries
        if end_row is None or end_col is None:
            print(f"⚠ Warning: Invalid table boundaries - end_row: {end_row}, end_col: {end_col}")
            return None
        

        return TableLocation(
            start_row=start_row,
            start_col=start_col,
            end_row=end_row,
            end_col=end_col,
            width_row=width_row,
            height_col=height_col,
            table_type="other",
            price_cols=None
        )
    
    def _is_valid_price_column(self, cell_value):
        """Check if column header contains valid keywords: Price, Not, Table"""
        if cell_value is None:
            return False
        
        cell_str = str(cell_value).lower().strip()
        valid_keywords = ["price", "not", "table"]
        return any(keyword in cell_str for keyword in valid_keywords)

    def extract_other_table_prices(self, sheet, table_loc: TableLocation, table_id: int, conn) -> int:
        """Extract prices from other-type table structure"""
        cursor = conn.cursor()
        price_count = 0
        
        # Detect if inch rows are separated or adjacent
        # Check if the next row after width_row is also an inch row
        next_row_value = sheet.cell(table_loc.width_row + 1, table_loc.start_col).value
        is_separated = not self.is_inch_value(next_row_value)
        
        # Get width cell
        for col in range(table_loc.height_col, table_loc.end_col + 1):
            # Check if column header contains valid keywords
            column_header = sheet.cell(table_loc.start_row, col).value
            if not self._is_valid_price_column(column_header):
                continue
                
            width = None
            
            # Get height cell - use appropriate step based on separation
            step = 2 if is_separated else 1
            end_row = table_loc.end_row + 1 if not is_separated else table_loc.end_row
            for row in range(table_loc.width_row, end_row, step):
                height = self.is_inch_value(sheet.cell(row, table_loc.start_col).value)
                if height is None:
                    continue
                
                # Get prices
                try:
                    # Normal price (inch row)
                    normal_price_cell = sheet.cell(row, col).value
                    normal_price = float(normal_price_cell) if normal_price_cell and isinstance(normal_price_cell, (int, float)) else None
                    
                    # Price with damper - adjust based on separation
                    if is_separated:
                        # Inch rows are separated by 1 row, damper price is in next row
                        damper_price_cell = sheet.cell(row + 1, col).value
                        damper_price = float(damper_price_cell) if damper_price_cell and isinstance(damper_price_cell, (int, float)) else None
                    else:
                        # Inch rows are adjacent, no damper price
                        damper_price = None
                    
                    # Insert if at least one price exists
                    if normal_price is not None or damper_price is not None:
                        cursor.execute('''
                            INSERT OR REPLACE INTO prices 
                            (table_id, width, height, normal_price, price_with_damper)
                            VALUES (?, ?, ?, ?, ?)
                        ''', (table_id, None, height, normal_price, damper_price))
                        price_count += 1
                except Exception:
                    continue
        
        return price_count
        