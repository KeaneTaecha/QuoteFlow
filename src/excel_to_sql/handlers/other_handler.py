"""
Other Table Handler Module
Handles other-type table detection and price extraction for the quotation system.
"""

import re
from typing import Optional, Tuple
from excel_to_sql.table_models import TableLocation


class OtherTableHandler:
    """Handles other-type table detection and price extraction"""
    
    def __init__(self, is_inch_value_func=None):
        self.is_inch_value = is_inch_value_func
    
    def _is_valid_width_value(self, cell_value, model_names=None):
        """Check if a cell value is a valid width (inch value or model name)"""
        if self.is_inch_value(cell_value):
            return True
        if model_names and cell_value is not None:
            cell_str = str(cell_value).strip()
            print(f"Checking if {cell_str} is in {model_names}")
            return cell_str and cell_str in model_names
        return False
    
    def _get_width_value(self, cell_value, model_names=None):
        """Get width value - returns integer for inch values, or None for model names"""
        inch_value = self.is_inch_value(cell_value)
        if inch_value is not None:
            return inch_value
        # For model names, we return None as width (model names are identifiers, not widths)
        if model_names and cell_value is not None:
            cell_str = str(cell_value).strip()
            if cell_str and cell_str in model_names:
                return None  # Model name found but not a numeric width
        return None
    
    def get_other_table_bounderies(self, sheet, start_row: int, start_col: int, model_names=None) -> Optional[TableLocation]:
        """Get boundaries of other-type table structure using keyword recognition"""
        
        # Look for width row (containing inch values like "4"", "6"" or model names)
        width_row = None
        for row in range(start_row, start_row + 5):
            cell_value = sheet.cell(row, start_col).value
            if self._is_valid_width_value(cell_value, model_names):
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
            if not self._is_valid_width_value(cell_value, model_names):
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
    
    def _is_price_per_feet_column(self, cell_value):
        """Check if column header is 'Price / 1 Ft.' or similar variations"""
        if cell_value is None:
            return False
        
        cell_str = str(cell_value).lower().strip()
        # Check for variations like "Price / 1 Ft.", "Price/1 Ft.", "Price per 1 Ft.", etc.
        price_per_feet_patterns = [
            "price / 1 ft",
            "price/1 ft",
            "price per 1 ft",
            "price per ft",
            "price/ft",
            "price per foot",
            "price/foot"
        ]
        return any(pattern in cell_str for pattern in price_per_feet_patterns)

    def extract_other_table_prices(self, sheet, table_loc: TableLocation, table_id: int, conn, model_names=None) -> int:
        """Extract prices from other-type table structure"""
        cursor = conn.cursor()
        price_count = 0
        
        # Detect if inch rows are separated or adjacent
        # Check if the next row after width_row is also a valid width value
        next_row_value = sheet.cell(table_loc.width_row + 1, table_loc.start_col).value
        is_separated = not self._is_valid_width_value(next_row_value, model_names)
        
        # First pass: identify price columns
        price_per_foot_col = None
        valid_price_cols = []
        
        for col in range(table_loc.height_col, table_loc.end_col + 1):
            column_header = sheet.cell(table_loc.start_row, col).value
            
            if self._is_price_per_feet_column(column_header):
                price_per_foot_col = col
            elif self._is_valid_price_column(column_header):
                valid_price_cols.append(col)
        
        # If no price columns found, return early
        if price_per_foot_col is None and not valid_price_cols:
            return 0
        
        # Second pass: collect prices for each width
        height = None
        step = 2 if is_separated else 1
        end_row = table_loc.end_row + 1 if not is_separated else table_loc.end_row
        
        for row in range(table_loc.width_row, end_row, step):
            cell_value = sheet.cell(row, table_loc.start_col).value
            width = self._get_width_value(cell_value, model_names)
            # If width column contains model name, width will be None (saved as NULL in database)
            
            # Collect all price values for this width
            normal_price = None
            damper_price = None
            price_per_foot = None
            
            try:
                # Get price per foot from price_per_foot_col if it exists
                if price_per_foot_col is not None:
                    cell_value = sheet.cell(row, price_per_foot_col).value
                    if cell_value and isinstance(cell_value, (int, float)):
                        price_per_foot = float(cell_value)
                
                # Get normal price and damper price from valid price columns
                # Use the first valid price column found
                if valid_price_cols:
                    col = valid_price_cols[0]  # Use first valid price column
                    cell_value = sheet.cell(row, col).value
                    if cell_value and isinstance(cell_value, (int, float)):
                        normal_price = float(cell_value)
                    
                    # Price with damper - adjust based on separation (only for valid price columns)
                    if is_separated:
                        damper_price_cell = sheet.cell(row + 1, col).value
                        if damper_price_cell and isinstance(damper_price_cell, (int, float)):
                            damper_price = float(damper_price_cell)
                
                # Insert if at least one price exists
                if normal_price is not None or damper_price is not None or price_per_foot is not None:
                    cursor.execute('''
                        INSERT OR REPLACE INTO prices 
                        (table_id, height, width, normal_price, price_with_damper, price_per_foot)
                        VALUES (?, ?, ?, ?, ?, ?)
                    ''', (table_id, None, width, normal_price, damper_price, price_per_foot))
                    price_count += 1
            except Exception:
                continue
        
        return price_count
        