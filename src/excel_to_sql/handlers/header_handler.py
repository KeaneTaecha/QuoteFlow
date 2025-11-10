"""
Header Table Handler Module
Handles Header sheet table detection and column identification using keyword recognition.
"""

from typing import Optional, Dict, List, Tuple
from table_models import TableLocation


class HeaderTableHandler:
    """Handles Header sheet table detection and column identification"""
    
    def __init__(self):
        # Keywords for column detection
        # Note: table_id is now auto-generated, not read from Excel
        self.column_keywords = {
            'sheet_name': ['sheet', 'sheet name', 'sheet_name', 'sheetname'],
            'model': ['model', 'models', 'product', 'product model'],
            'tb_modifier': ['tb modifier', 'tb_modifier', 'tb modifier equation', 'tb_modifier_equation', 'base price modifier', 'base_price_modifier', 'bp modifier', 'bp_modifier'],
            'anodized': ['anodized', 'aluminum', 'anodized aluminum', 'anodized multiplier', 'anodized multiplier', 'anodized aluminum multiplier'],
            'powder_coated': ['powder coated', 'powder_coated', 'powder', 'coated', 'powder coated multiplier', 'powdercoated'],
            'wd': ['wd', 'with damper', 'with_damper', 'wd multiplier', 'wd equation', 'damper', 'damper multiplier', 'damper equation']
        }
    
    def detect_header_table(self, sheet) -> Optional[TableLocation]:
        """Detect the Header table location and return its boundaries"""
        
        # Look for the header row containing keywords
        header_row = self._find_header_row(sheet)
        if header_row is None:
            print("⚠ Warning: No header row found in Header sheet")
            return None
        
        # Find the data start row (first row after header with actual data)
        data_start_row = self._find_data_start_row(sheet, header_row)
        if data_start_row is None:
            print("⚠ Warning: No data rows found in Header sheet")
            return None
        
        # Find the end of data
        data_end_row = self._find_data_end_row(sheet, data_start_row)
        if data_end_row is None:
            print("⚠ Warning: No valid data end found in Header sheet")
            return None
        
        # Find the extent of columns
        start_col, end_col = self._find_column_bounds(sheet, header_row)
        if start_col is None or end_col is None:
            print("⚠ Warning: Could not determine column bounds in Header sheet")
            return None
        
        
        return TableLocation(
            start_row=header_row,
            start_col=start_col,
            end_row=data_end_row,
            end_col=end_col,
            width_row=header_row,  # Header row contains column names
            height_col=start_col,  # First column
            table_type="header",
            price_cols=None
        )
    
    def _find_header_row(self, sheet) -> Optional[int]:
        """Find the row containing header keywords"""
        
        for row in range(1, min(20, sheet.max_row + 1)):  # Search first 20 rows
            row_keywords = []
            for col in range(1, min(20, sheet.max_column + 1)):  # Search first 20 columns
                cell_value = sheet.cell(row, col).value
                if cell_value is None:
                    continue
                
                cell_str = str(cell_value).lower().strip()
                
                # Check if this cell contains any header keywords
                for field_name, keywords in self.column_keywords.items():
                    if any(keyword in cell_str for keyword in keywords):
                        row_keywords.append(field_name)
            
            # If we found at least 2 different keyword types, this is likely the header row
            if len(set(row_keywords)) >= 2:
                return row
        
        return None
    
    def _find_data_start_row(self, sheet, header_row: int) -> Optional[int]:
        """Find the first row with actual data after the header"""
        for row in range(header_row + 1, min(header_row + 50, sheet.max_row + 1)):
            # Check if this row has any non-empty cells
            for col in range(1, min(20, sheet.max_column + 1)):
                cell_value = sheet.cell(row, col).value
                if cell_value is not None and str(cell_value).strip():
                    return row
        
        return None
    
    def _find_data_end_row(self, sheet, data_start_row: int) -> Optional[int]:
        """Find the last row with data"""
        last_data_row = None
        
        for row in range(data_start_row, min(data_start_row + 100, sheet.max_row + 1)):
            # Check if this row has any non-empty cells
            has_data = False
            for col in range(1, min(20, sheet.max_column + 1)):
                cell_value = sheet.cell(row, col).value
                if cell_value is not None and str(cell_value).strip():
                    has_data = True
                    break
            
            if has_data:
                last_data_row = row
            else:
                # If we hit an empty row, check a few more rows to be sure
                empty_count = 0
                for check_row in range(row, min(row + 3, sheet.max_row + 1)):
                    row_empty = True
                    for col in range(1, min(20, sheet.max_column + 1)):
                        cell_value = sheet.cell(check_row, col).value
                        if cell_value is not None and str(cell_value).strip():
                            row_empty = False
                            break
                    if row_empty:
                        empty_count += 1
                
                if empty_count >= 2:  # Two consecutive empty rows, we're done
                    break
        
        if last_data_row:
            return last_data_row
        
        return None
    
    def _find_column_bounds(self, sheet, header_row: int) -> Tuple[Optional[int], Optional[int]]:
        """Find the start and end columns of the table"""
        start_col = None
        end_col = None
        
        # Find start column (first non-empty cell in header row)
        for col in range(1, min(20, sheet.max_column + 1)):
            cell_value = sheet.cell(header_row, col).value
            if cell_value is not None and str(cell_value).strip():
                start_col = col
                break
        
        # Find end column (last non-empty cell in header row)
        for col in range(sheet.max_column, 0, -1):
            cell_value = sheet.cell(header_row, col).value
            if cell_value is not None and str(cell_value).strip():
                end_col = col
                break
        
        if start_col and end_col:
            pass
        
        return start_col, end_col
    
    def get_column_mapping(self, sheet, table_loc: TableLocation) -> Dict[str, int]:
        """Get column indices for each field based on keywords"""
        column_mapping = {}
        
        # Search through the header row to find column indices
        for col in range(table_loc.start_col, table_loc.end_col + 1):
            cell_value = sheet.cell(table_loc.start_row, col).value
            if cell_value is None:
                continue
            
            cell_str = str(cell_value).lower().strip()
            
            # Check each keyword group
            for field_name, keywords in self.column_keywords.items():
                if any(keyword in cell_str for keyword in keywords):
                    column_mapping[field_name] = col
                    break
        
        return column_mapping
    
    def extract_header_data(self, sheet, table_loc: TableLocation, column_mapping: Dict[str, int]) -> List[Dict]:
        """Extract header data using the column mapping"""
        header_data = []
        
        # Auto-generate table IDs starting from 1
        table_id_counter = 1
        
        # Extract data from each row
        for row in range(table_loc.start_row + 1, table_loc.end_row + 1):
            # Get values using column mapping
            # Note: table_id is now auto-generated, not read from Excel
            sheet_name = self._get_cell_value(sheet, row, column_mapping.get('sheet_name'))
            model = self._get_cell_value(sheet, row, column_mapping.get('model'))
            tb_modifier = self._get_cell_value(sheet, row, column_mapping.get('tb_modifier'))
            anodized_multiplier = self._get_cell_value(sheet, row, column_mapping.get('anodized'))
            powder_coated_multiplier = self._get_cell_value(sheet, row, column_mapping.get('powder_coated'))
            wd_multiplier = self._get_cell_value(sheet, row, column_mapping.get('wd'))
            
            # Skip rows with missing essential data (table_id is no longer required from Excel)
            if sheet_name is None or model is None:
                continue
            
            # Parse models (comma-separated)
            models = [m.strip() for m in str(model).split(',')]
            
            # Parse TB modifier (can be number or equation)
            tb_modifiers = self._parse_multipliers(tb_modifier)
            
            # Parse multipliers (can be numbers or equations)
            anodized_multipliers = self._parse_multipliers(anodized_multiplier)
            powder_coated_multipliers = self._parse_multipliers(powder_coated_multiplier)
            wd_multipliers = self._parse_multipliers(wd_multiplier)
            
            # Adjust TB modifier list to match model count
            if len(tb_modifiers) == 1:
                tb_modifiers = tb_modifiers * len(models)
            elif len(tb_modifiers) > 0 and len(tb_modifiers) < len(models):
                tb_modifiers.extend([tb_modifiers[-1]] * (len(models) - len(tb_modifiers)))
            
            # Adjust multiplier lists to match model count
            if len(anodized_multipliers) == 1:
                anodized_multipliers = anodized_multipliers * len(models)
            elif len(anodized_multipliers) > 0 and len(anodized_multipliers) < len(models):
                anodized_multipliers.extend([anodized_multipliers[-1]] * (len(models) - len(anodized_multipliers)))
            
            if len(powder_coated_multipliers) == 1:
                powder_coated_multipliers = powder_coated_multipliers * len(models)
            elif len(powder_coated_multipliers) > 0 and len(powder_coated_multipliers) < len(models):
                powder_coated_multipliers.extend([powder_coated_multipliers[-1]] * (len(models) - len(powder_coated_multipliers)))
            
            if len(wd_multipliers) == 1:
                wd_multipliers = wd_multipliers * len(models)
            elif len(wd_multipliers) > 0 and len(wd_multipliers) < len(models):
                wd_multipliers.extend([wd_multipliers[-1]] * (len(models) - len(wd_multipliers)))
            
            entry = {
                'table_id': table_id_counter,  # Auto-generated sequential ID
                'sheet_name': str(sheet_name),
                'models': models,
                'tb_modifiers': tb_modifiers,
                'anodized_multipliers': anodized_multipliers,
                'powder_coated_multipliers': powder_coated_multipliers,
                'wd_multipliers': wd_multipliers
            }
            
            header_data.append(entry)
            table_id_counter += 1  # Increment for next table
        
        return header_data
    
    def _get_cell_value(self, sheet, row: int, col: Optional[int]) -> Optional[str]:
        """Get cell value safely"""
        if col is None:
            return None
        
        cell_value = sheet.cell(row, col).value
        if cell_value is None:
            return None
        
        return str(cell_value).strip()
    
    def _parse_multipliers(self, multiplier_value) -> List[Optional[str]]:
        """Parse multiplier values (can be numbers or equations)"""
        multipliers = []
        
        if multiplier_value is not None and str(multiplier_value).strip().lower() != 'none':
            for m in str(multiplier_value).split(','):
                m = m.strip()
                if m.lower() != 'none' and m != '':
                    multipliers.append(m)  # Store as string (number or equation)
                else:
                    multipliers.append(None)
        
        return multipliers
