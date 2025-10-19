"""
Default Table Handler Module
Handles default/standard table detection and price extraction for the quotation system.
"""

from typing import Optional
from table_models import TableLocation


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
            print(f"        Debug: Invalid table boundaries - end_row: {end_row}, end_col: {end_col}")
            return None
        
        print(f"        Debug: Start_row: ({start_row})")
        print(f"        Debug: Start_col: ({start_col})")
        print(f"        Debug: End_row: ({end_row})")
        print(f"        Debug: End_col: ({end_col})")
        print(f"        Debug: Width_row: ({width_row})")
        print(f"        Debug: Height_col: ({height_col})")

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
        
        # Get width cell
        for col in range(table_loc.height_col, table_loc.end_col + 1):
            width = self.is_inch_value(sheet.cell(table_loc.start_row, col).value)
            if width is None:
                continue
            
            # Get height cell
            for row in range(table_loc.width_row, table_loc.end_row, 2):
                height = self.is_inch_value(sheet.cell(row, table_loc.start_col).value)
                if height is None:
                    continue
                
                # Get prices
                try:
                    # Normal price (inch row)
                    normal_price_cell = sheet.cell(row, col).value
                    normal_price = float(normal_price_cell) if normal_price_cell and isinstance(normal_price_cell, (int, float)) else None
                    
                    # Price with damper (mm row - next row)
                    damper_price_cell = sheet.cell(row + 1, col).value
                    damper_price = float(damper_price_cell) if damper_price_cell and isinstance(damper_price_cell, (int, float)) else None
                    
                    # Insert if at least one price exists
                    if normal_price is not None or damper_price is not None:
                        cursor.execute('''
                            INSERT OR REPLACE INTO prices 
                            (table_id, width, height, normal_price, price_with_damper)
                            VALUES (?, ?, ?, ?, ?)
                        ''', (table_id, width, height, normal_price, damper_price))
                        price_count += 1
                except Exception:
                    continue
        
        return price_count
