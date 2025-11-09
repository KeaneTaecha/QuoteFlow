"""
Default Table Handler Module
Handles default/standard table detection and price extraction for the quotation system.
"""

from typing import Optional
from excel_to_sql.table_models import TableLocation


class DefaultTableHandler:
    """Handles default/standard table detection and price extraction"""
    
    def __init__(self, is_inch_value_func=None):
        self.is_inch_value = is_inch_value_func
    
    def get_default_table_bounderies(self, sheet, start_row: int, start_col: int) -> Optional[TableLocation]:
        """Find the boundaries of a default/standard table starting from a given position"""
        
        # Look for width row (containing inch values like "4"", "6"")
        width_row = None
        for row in range(start_row, start_row + 5):
            cell_value = sheet.cell(row, start_col).value
            if self.is_inch_value(cell_value):
                width_row = row
                break
        
        if width_row is None:
            return None
        
        # Find height column
        height_col = None
        for col in range(start_col, start_col + 5):
            cell_value = sheet.cell(start_row, col).value
            if self.is_inch_value(cell_value):
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
            if not self.is_inch_value(cell_value):
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
            table_type="standard",
            price_cols=None
        )

    def extract_default_table_prices(self, sheet, table_loc: TableLocation, table_id: int, conn) -> int:
        """Extract price data from a default/standard table and store in database"""
        cursor = conn.cursor()
        price_count = 0
        
        # Detect if inch rows are separated or adjacent
        # Check if the next row after width_row is also an inch row
        next_row_value = sheet.cell(table_loc.width_row + 1, table_loc.start_col).value
        is_separated = not self.is_inch_value(next_row_value)
        
        # Get height cell (from header row)
        for col in range(table_loc.height_col, table_loc.end_col + 1):
            height = self.is_inch_value(sheet.cell(table_loc.start_row, col).value)
            if height is None:
                continue
            
            # Get width cell (from first column) - use appropriate step based on separation
            step = 2 if is_separated else 1
            end_row = table_loc.end_row + 1 if not is_separated else table_loc.end_row
            for row in range(table_loc.width_row, end_row, step):
                width = self.is_inch_value(sheet.cell(row, table_loc.start_col).value)
                if width is None:
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
                            (table_id, height, width, normal_price, price_with_damper)
                            VALUES (?, ?, ?, ?, ?)
                        ''', (table_id, height, width, normal_price, damper_price))
                        price_count += 1
                except Exception:
                    continue
        
        # Extract multipliers from table boundaries
        self._extract_table_multipliers(sheet, table_loc, table_id, conn)
        
        return price_count
    
    def _extract_table_multipliers(self, sheet, table_loc: TableLocation, table_id: int, conn):
        """Extract multipliers for each width row and height column"""
        cursor = conn.cursor()
        
        try:
            # Detect if inch rows are separated or adjacent
            next_row_value = sheet.cell(table_loc.width_row + 1, table_loc.start_col).value
            is_separated = not self.is_inch_value(next_row_value)
            step = 2 if is_separated else 1
            end_row = table_loc.end_row + 1 if not is_separated else table_loc.end_row
            
            # Extract multipliers for each width row (height exceeded multipliers - regular and WD)
            for row in range(table_loc.width_row, end_row, step):
                width = self.is_inch_value(sheet.cell(row, table_loc.start_col).value)
                if width is None:
                    continue
                
                # Check for regular height exceeded multiplier in the column to the right of the table
                height_mult_cell = sheet.cell(row, table_loc.end_col + 1).value
                height_multiplier = None
                if height_mult_cell is not None and isinstance(height_mult_cell, (int, float)):
                    height_multiplier = float(height_mult_cell)
                
                # Check for WD height exceeded multiplier (1 cell below regular multiplier)
                height_mult_wd_cell = sheet.cell(row + 1, table_loc.end_col + 1).value
                height_multiplier_wd = None
                if height_mult_wd_cell is not None and isinstance(height_mult_wd_cell, (int, float)):
                    height_multiplier_wd = float(height_mult_wd_cell)
                
                if height_multiplier is not None or height_multiplier_wd is not None:
                    cursor.execute('''
                        INSERT OR REPLACE INTO row_multipliers 
                        (table_id, width, height_exceeded_multiplier, height_exceeded_multiplier_wd)
                        VALUES (?, ?, ?, ?)
                    ''', (table_id, width, height_multiplier, height_multiplier_wd))
            
            # Extract multipliers for each height column (width exceeded multipliers - regular and WD)
            for col in range(table_loc.height_col, table_loc.end_col + 1):
                height = self.is_inch_value(sheet.cell(table_loc.start_row, col).value)
                if height is None:
                    continue
                
                # Check for regular width exceeded multiplier in the row below the table
                width_mult_cell = sheet.cell(table_loc.end_row + 1, col).value
                width_multiplier = None
                if width_mult_cell is not None and isinstance(width_mult_cell, (int, float)):
                    width_multiplier = float(width_mult_cell)
                
                # Check for WD width exceeded multiplier (1 cell below regular multiplier)
                width_mult_wd_cell = sheet.cell(table_loc.end_row + 2, col).value
                width_multiplier_wd = None
                if width_mult_wd_cell is not None and isinstance(width_mult_wd_cell, (int, float)):
                    width_multiplier_wd = float(width_mult_wd_cell)
                
                if width_multiplier is not None or width_multiplier_wd is not None:
                    cursor.execute('''
                        INSERT OR REPLACE INTO column_multipliers 
                        (table_id, height, width_exceeded_multiplier, width_exceeded_multiplier_wd)
                        VALUES (?, ?, ?, ?)
                    ''', (table_id, height, width_multiplier, width_multiplier_wd))
            
            
        except Exception as e:
            print(f"⚠ Warning: Error extracting multipliers for table {table_id}: {e}")
