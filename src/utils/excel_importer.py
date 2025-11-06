"""
Excel Importer Module
Handles importing items from Excel files into quotations.
"""

import re
import openpyxl
from typing import List, Dict, Optional, Tuple
from sql.price_loader import PriceListLoader


class ExcelItemImporter:
    """Handles importing items from Excel files"""
    
    def __init__(self, price_loader: PriceListLoader, available_models: List[str]):
        """
        Initialize the Excel importer
        
        Args:
            price_loader: PriceListLoader instance for database access
            available_models: List of available product models
        """
        self.price_loader = price_loader
        self.available_models = available_models
    
    def parse_excel_file(self, file_path: str, progress_callback=None) -> List[Dict]:
        """
        Parse Excel file and extract items based on Model, Detail, Width, Height, Unit, Quantity, Finish columns
        
        Args:
            file_path: Path to the Excel file
            progress_callback: Optional callback function(progress_percent_0_to_100, status_text) to report progress
                               The callback receives internal progress 0-100%, caller should map to overall progress
        """
        if progress_callback:
            progress_callback(5, 'Loading Excel file...')
        
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        sheet = wb.active  # Use the first/active sheet
        
        if progress_callback:
            progress_callback(10, 'Searching for header row...')
        
        # Find header row by searching for keywords
        header_row = None
        column_mapping = {}
        
        # Search for header row (search first 20 rows)
        for row in range(1, min(21, sheet.max_row + 1)):
            # Check all columns in this row for header keywords
            row_mapping = {}
            for col in range(1, sheet.max_column + 1):
                cell_value = sheet.cell(row, col).value
                if cell_value is None:
                    continue
                
                cell_str = str(cell_value).strip().lower()
                
                # Check for keywords
                if 'model' in cell_str and 'model' not in row_mapping:
                    row_mapping['model'] = col
                elif 'detail' in cell_str and 'detail' not in row_mapping:
                    row_mapping['detail'] = col
                elif 'width' in cell_str and 'width' not in row_mapping:
                    row_mapping['width'] = col
                elif 'height' in cell_str and 'height' not in row_mapping:
                    row_mapping['height'] = col
                elif 'unit' in cell_str and 'unit' not in row_mapping:
                    row_mapping['unit'] = col
                elif 'quantity' in cell_str and 'quantity' not in row_mapping:
                    row_mapping['quantity'] = col
                elif 'finish' in cell_str and 'finish' not in row_mapping:
                    row_mapping['finish'] = col
            
            # If we found Model column and at least a few other columns, use this row
            if 'model' in row_mapping and len(row_mapping) >= 2:
                header_row = row
                column_mapping = row_mapping
                break
        
        if header_row is None or 'model' not in column_mapping:
            raise ValueError("Could not find required 'Model' column in Excel file")
        
        # Extract items from rows below header
        items = []
        model_col = column_mapping.get('model')
        detail_col = column_mapping.get('detail')
        width_col = column_mapping.get('width')
        height_col = column_mapping.get('height')
        unit_col = column_mapping.get('unit')
        quantity_col = column_mapping.get('quantity')
        finish_col = column_mapping.get('finish')
        
        # Calculate total rows to process
        total_rows = sheet.max_row - header_row
        processed_rows = 0
        
        # Process rows below header
        for row in range(header_row + 1, sheet.max_row + 1):
            model_value = self._get_cell_value(sheet, row, model_col)
            
            # Skip empty rows
            if model_value is None or str(model_value).strip() == '':
                processed_rows += 1
                # Update progress every 10 rows or at the end
                if progress_callback and total_rows > 0 and (processed_rows % 10 == 0 or processed_rows == total_rows):
                    progress = 15 + int((processed_rows / total_rows) * 85)  # 15% to 100% internal
                    progress = min(progress, 100)
                    progress_callback(progress, f'Reading row {row} of {sheet.max_row}...')
                continue
            
            model_str = str(model_value).strip()
            
            # Get other column values
            detail_value = self._get_cell_value(sheet, row, detail_col) if detail_col else None
            width_value = self._get_cell_value(sheet, row, width_col) if width_col else None
            height_value = self._get_cell_value(sheet, row, height_col) if height_col else None
            unit_value = self._get_cell_value(sheet, row, unit_col) if unit_col else None
            quantity_value = self._get_cell_value(sheet, row, quantity_col) if quantity_col else None
            finish_value = self._get_cell_value(sheet, row, finish_col) if finish_col else None
            
            # Check if this is a title (Model has text but other columns are empty)
            has_detail = detail_value is not None and str(detail_value).strip() != ''
            has_width = width_value is not None and str(width_value).strip() != ''
            has_height = height_value is not None and str(height_value).strip() != ''
            has_quantity = quantity_value is not None and str(quantity_value).strip() != ''
            has_finish = finish_value is not None and str(finish_value).strip() != ''
            
            if not has_detail and not has_width and not has_height and not has_quantity and not has_finish:
                # This is a title
                items.append({
                    'is_title': True,
                    'title': model_str,
                    'product_code': '',
                    'size': '',
                    'finish': '',
                    'quantity': 0,
                    'unit_price': 0,
                    'discount': 0,
                    'discounted_unit_price': 0,
                    'total': 0,
                    'rounded_size': None,
                    'detail': ''
                })
            else:
                # This is a product item
                item = {
                    'model': model_str,
                    'detail': str(detail_value).strip() if detail_value else '',
                    'width': width_value,
                    'height': height_value,
                    'unit': str(unit_value).strip().lower() if unit_value else 'inches',
                    'quantity': self._parse_number(quantity_value) if quantity_value else 1,
                    'finish': str(finish_value).strip() if finish_value else None
                }
                items.append(item)
            
            processed_rows += 1
            # Update progress every 10 rows or at the end
            if progress_callback and total_rows > 0 and (processed_rows % 10 == 0 or processed_rows == total_rows):
                progress = 15 + int((processed_rows / total_rows) * 85)  # 15% to 100% internal
                progress = min(progress, 100)
                progress_callback(progress, f'Reading row {row} of {sheet.max_row}...')
        
        if progress_callback:
            progress_callback(100, 'Parsing complete!')
        
        wb.close()
        return items
    
    def _get_cell_value(self, sheet, row, col):
        """Get cell value safely, returning None if cell doesn't exist"""
        if col is None:
            return None
        try:
            return sheet.cell(row, col).value
        except:
            return None
    
    def _parse_number(self, value):
        """Parse a number from a cell value"""
        if value is None:
            return 1
        try:
            # Try to convert to float first, then int
            num = float(str(value).strip())
            return int(num) if num.is_integer() else num
        except:
            return 1
    
    def _parse_dimension_with_unit(self, value, default_unit):
        """Parse dimension value, handling quote (") for inches"""
        if value is None:
            return None, None
        
        value_str = str(value).strip()
        
        # Check if there's a quote (") anywhere in the string (e.g., "3"" or "3 inches" or just "3")
        if '"' in value_str:
            # Extract number and treat as inches
            match = re.search(r'(\d+(?:\.\d+)?)', value_str)
            if match:
                num_value = float(match.group(1))
                return num_value, 'inches'
        
        # Try to extract number
        match = re.search(r'(\d+(?:\.\d+)?)', value_str)
        if match:
            num_value = float(match.group(1))
            return num_value, default_unit
        
        return None, None
    
    def _create_invalid_item(self, item_data, error_message):
        """Create an invalid item entry for display in the quote"""
        model = item_data.get('model', 'Unknown').strip()
        detail = item_data.get('detail', '').strip()
        width_value = item_data.get('width')
        height_value = item_data.get('height')
        quantity = item_data.get('quantity', 1)
        finish_value = item_data.get('finish')
        
        # Create size string from available dimensions
        size_parts = []
        if width_value is not None:
            size_parts.append(f"W: {width_value}")
        if height_value is not None:
            size_parts.append(f"H: {height_value}")
        size_str = " x ".join(size_parts) if size_parts else "N/A"
        
        # Create finish string
        finish_str = 'INVALID'
        if finish_value:
            finish_str = f"INVALID ({finish_value})"
        
        return {
            'is_invalid': True,
            'product_code': model,
            'size': size_str,
            'finish': finish_str,
            'quantity': quantity,
            'unit_price': 0,
            'discount': 0,
            'discounted_unit_price': 0,
            'total': 0,
            'rounded_size': None,
            'detail': detail,
            'error_message': error_message
        }
    
    def add_item_from_excel(self, item_data):
        """Add an item from Excel data to the quote. Returns dict with 'success' and 'error' keys."""
        model = item_data.get('model', '').strip()
        detail = item_data.get('detail', '').strip()
        width_value = item_data.get('width')
        height_value = item_data.get('height')
        unit_str = item_data.get('unit', 'inches').lower()
        quantity = item_data.get('quantity', 1)
        finish_from_excel = item_data.get('finish')  # Finish from Excel file
        
        # Determine default unit (inches or millimeters)
        default_unit = 'inches'
        if 'mm' in unit_str or 'millimeter' in unit_str:
            default_unit = 'millimeters'
        
        # Parse width and height with unit handling
        width, width_unit = self._parse_dimension_with_unit(width_value, default_unit)
        height, height_unit = self._parse_dimension_with_unit(height_value, default_unit)
        
        if width is None and height is None:
            # Check if this is an "other table" product (diameter-based)
            # For now, we'll skip items without dimensions
            return {'success': False, 'error': 'Missing dimensions (width and height)'}
        
        # Check if product has filter suffix (+F.xxx)
        filter_type = None
        base_model = model
        if '+F.' in model:
            parts = model.split('+F.', 1)
            if len(parts) == 2:
                base_model = parts[0].strip()
                filter_type = parts[1].strip()
        
        # Check if product exists in database
        product = base_model
        has_wd = False
        if product.endswith("(WD)"):
            product = product[:-4].strip()
            has_wd = True
        
        # Validate product exists
        if not self.available_models:
            return {'success': False, 'error': 'Price database not initialized'}
        
        # Check if product exists (with or without WD)
        product_exists = False
        if product in self.available_models:
            product_exists = True
        elif has_wd and product in [m.replace("(WD)", "").strip() for m in self.available_models]:
            product_exists = True
        
        if not product_exists:
            # Try to find a close match
            matching_models = [m for m in self.available_models if product.lower() in m.lower() or m.lower() in product.lower()]
            if matching_models:
                product = matching_models[0].replace("(WD)", "").strip()
                if matching_models[0].endswith("(WD)"):
                    has_wd = True
            else:
                return {'success': False, 'error': f'Product "{model}" not found in database'}
        
        # Get available finishes
        finishes = self.price_loader.get_available_finishes(product)
        if not finishes:
            return {'success': False, 'error': f'No finishes available for product "{product}"'}
        
        # Determine which finish to use
        finish = self._match_finish(finish_from_excel, finishes)
        if finish is None:
            return {'success': False, 'error': f'Finish "{finish_from_excel}" not available for product "{product}". Available finishes: {", ".join(finishes)}'}
        
        # Check if this is an "other table" product (diameter-based)
        is_other_table = self.price_loader.is_other_table(product)
        
        if is_other_table:
            # Handle diameter-based products
            size = width if width is not None else (height if height is not None else 4)
            size_unit = width_unit if width is not None else (height_unit if height is not None else 'inches')
            
            # Convert to inches if needed
            if size_unit == 'millimeters':
                size_inches = size / 25.4
            else:
                size_inches = size
            
            # Find rounded size and get price
            rounded_size = self.price_loader.find_rounded_other_table_size(product, finish, size_inches)
            if not rounded_size:
                return {'success': False, 'error': f'Size not available for {product}'}
            
            unit_price = self.price_loader.get_price_for_other_table(product, finish, rounded_size, has_wd)
            if unit_price == 0:
                return {'success': False, 'error': f'Price not available for {product} {rounded_size}'}
            
            # Add filter price if filter is specified
            if filter_type:
                filter_price = self._get_filter_price(filter_type, size_inches)
                if filter_price is None:
                    return {'success': False, 'error': f'Filter "{filter_type}" not found in database'}
                unit_price += filter_price
            
            # Store original size
            if size_unit == 'millimeters':
                original_size = f"{size}mm"
            else:
                original_size = f'{size}"'
            
            # Build product code with filter if applicable
            product_code = f"{product}(WD)" if has_wd else product
            if filter_type:
                product_code = f"{product_code}+F.{filter_type}"
            
            quote_item = {
                'product_code': product_code,
                'size': original_size,
                'finish': finish,
                'quantity': quantity,
                'unit_price': unit_price,
                'discount': 0,
                'discounted_unit_price': unit_price,
                'total': unit_price * quantity,
                'rounded_size': rounded_size,
                'detail': detail
            }
        else:
            # Handle width/height-based products
            if width is None or height is None:
                return {'success': False, 'error': f'Width and height required for {product}'}
            
            # Convert to inches if needed
            if width_unit == 'millimeters':
                width_inches = width / 25.4
            else:
                width_inches = width
            
            if height_unit == 'millimeters':
                height_inches = height / 25.4
            else:
                height_inches = height
            
            # Validate dimensions
            if height_inches > width_inches:
                return {'success': False, 'error': f'Height must be less than width for {product}'}
            
            # Find rounded size and get price
            rounded_size = self.price_loader.find_rounded_default_table_size(product, finish, width_inches, height_inches)
            if not rounded_size:
                # Try direct lookup for exceeded dimensions
                unit_price = self.price_loader.get_price_for_default_table(product, finish, f'{width_inches}" x {height_inches}"', has_wd)
                if unit_price == 0:
                    return {'success': False, 'error': f'Size not available for {product}'}
                rounded_size = f'{width_inches}" x {height_inches}"'
            else:
                unit_price = self.price_loader.get_price_for_default_table(product, finish, rounded_size, has_wd)
                if unit_price == 0:
                    return {'success': False, 'error': f'Price not available for {product} {rounded_size}'}
            
            # Add filter price if filter is specified
            if filter_type:
                # Pass actual width and height to filter (filter will use max for sizing)
                filter_price = self._get_filter_price(filter_type, max(width_inches, height_inches), width_inches, height_inches)
                if filter_price is None:
                    return {'success': False, 'error': f'Filter "{filter_type}" not found in database'}
                unit_price += filter_price
            
            # Store original size
            if width_unit == 'millimeters' and height_unit == 'millimeters':
                original_size = f"{width}mm x {height}mm"
            elif width_unit == 'inches' and height_unit == 'inches':
                original_size = f'{width}" x {height}"'
            else:
                # Mixed units
                width_str = f"{width}mm" if width_unit == 'millimeters' else f'{width}"'
                height_str = f"{height}mm" if height_unit == 'millimeters' else f'{height}"'
                original_size = f'{width_str} x {height_str}'
            
            # Build product code with filter if applicable
            product_code = f"{product}(WD)" if has_wd else product
            if filter_type:
                product_code = f"{product_code}+F.{filter_type}"
            
            quote_item = {
                'product_code': product_code,
                'size': original_size,
                'finish': finish,
                'quantity': quantity,
                'unit_price': unit_price,
                'discount': 0,
                'discounted_unit_price': unit_price,
                'total': unit_price * quantity,
                'rounded_size': rounded_size,
                'detail': detail
            }
        
        return {'success': True, 'item': quote_item, 'error': None}
    
    def _match_finish(self, finish_from_excel, finishes):
        """Match finish from Excel with available finishes"""
        if not finish_from_excel:
            return finishes[0] if finishes else None
        
        finish_str = str(finish_from_excel).strip()
        finish_lower = finish_str.lower()
        
        # Define keywords for each finish type
        finish_keywords = {
            'Anodized Aluminum': ['anodized', 'aluminum', 'anodized aluminum', 'anodised', 'aluminium', 'anodised aluminium'],
            'Powder Coated': ['powder', 'coated', 'powder coated', 'powdercoated', 'pc', 'powder coat', 'ขาวนวล', 'ขาวเงา', 'ขาวด้าน', 'ขาวบริสุทธิ์', 'ดำเงา', 'ดำด้าน', 'บรอนส์'],
            'Special Color': ['special', 'color', 'special color', 'custom', 'custom color', 'sc', 'special colour', 'custom colour']
        }
        
        # Thai color names that should be treated as color names (not removed)
        thai_colors = ['ขาวนวล', 'ขาวเงา', 'ขาวด้าน', 'ขาวบริสุทธิ์', 'ดำเงา', 'ดำด้าน', 'บรอนส์']
        
        # Try exact match first
        for available_finish in finishes:
            if available_finish.lower() == finish_lower:
                return available_finish
        
        # Try keyword matching
        finish = None
        # Check for Powder Coated with color
        powder_keywords = finish_keywords.get('Powder Coated', [])
        if any(keyword in finish_lower for keyword in powder_keywords):
            if 'Powder Coated' in finishes:
                # Extract color if provided (format: "Powder Coated - Color" or "Powder - Color")
                if ' - ' in finish_str:
                    color = finish_str.split(' - ', 1)[1].strip()
                    return f"Powder Coated - {color}"
                else:
                    # Extract color name by removing powder coated keywords from the string
                    color_name = finish_str.strip()
                    
                    # Check if the finish string is exactly a Thai color name
                    if color_name in thai_colors:
                        return f"Powder Coated - {color_name}"
                    else:
                        # Remove all powder coated keywords to get the color name
                        english_keywords = [k for k in powder_keywords if k not in thai_colors]
                        for keyword in sorted(english_keywords, key=len, reverse=True):
                            keyword_pattern = re.escape(keyword)
                            color_name = re.sub(keyword_pattern, '', color_name, flags=re.IGNORECASE).strip()
                        
                        color_name = re.sub(r'^[-,\s]+|[-,\s]+$', '', color_name).strip()
                        
                        if color_name:
                            return f"Powder Coated - {color_name}"
                        else:
                            return 'Powder Coated'
        
        # Check for Special Color with name
        if finish is None:
            special_keywords = finish_keywords.get('Special Color', [])
            if any(keyword in finish_lower for keyword in special_keywords):
                if 'Special Color' in finishes:
                    if ' - ' in finish_str:
                        color_name = finish_str.split(' - ', 1)[1].strip()
                        return f"Special Color - {color_name}"
                    else:
                        color_name = finish_str.strip()
                        for keyword in sorted(special_keywords, key=len, reverse=True):
                            keyword_pattern = re.escape(keyword)
                            color_name = re.sub(keyword_pattern, '', color_name, flags=re.IGNORECASE).strip()
                        
                        color_name = re.sub(r'^[-,\s]+|[-,\s]+$', '', color_name).strip()
                        
                        if color_name:
                            return f"Special Color - {color_name}"
                        else:
                            return 'Special Color'
        
        # Check for Anodized Aluminum
        if finish is None:
            anodized_keywords = finish_keywords.get('Anodized Aluminum', [])
            if any(keyword in finish_lower for keyword in anodized_keywords):
                if 'Anodized Aluminum' in finishes:
                    return 'Anodized Aluminum'
        
        # If still no match, try substring matching with available finishes
        if finish is None:
            for available_finish in finishes:
                available_lower = available_finish.lower()
                finish_words = finish_lower.split()
                available_words = available_lower.split()
                
                if any(word in available_lower for word in finish_words if len(word) > 2):
                    return available_finish
                
                if any(word in finish_lower for word in available_words if len(word) > 2):
                    return available_finish
        
        return finish if finish else finishes[0] if finishes else None
    
    def _get_filter_price(self, filter_type, size_inches, width_inches=None, height_inches=None):
        """
        Find filter product in database and get its price.
        
        Args:
            filter_type: The filter type (e.g., "Nylon")
            size_inches: The size in inches (for diameter-based filters) or max dimension
            width_inches: Optional width in inches for dimension-based filters
            height_inches: Optional height in inches for dimension-based filters
            
        Returns:
            Filter price or None if not found
        """
        if not self.price_loader:
            return None
        
        # Get all available models
        all_models = self.price_loader.get_available_models()
        
        # Search for filter product that matches the filter type
        filter_type_lower = filter_type.lower()
        matching_filter = None
        
        # Simple matching: find any product that contains the filter type word
        for model in all_models:
            model_lower = model.lower()
            if filter_type_lower in model_lower:
                matching_filter = model
                break
        
        if not matching_filter:
            return None
        
        # Get available finishes for the filter
        finishes = self.price_loader.get_available_finishes(matching_filter)
        if not finishes:
            return None
        
        # Use first available finish
        finish = finishes[0]
        
        # Check if filter is diameter-based (other table) or dimension-based
        is_other_table = self.price_loader.is_other_table(matching_filter)
        
        if is_other_table:
            # For diameter-based filters, get base price directly from database
            price = self._get_base_price_for_other_table(matching_filter, size_inches)
            return price if price and price > 0 else None
        else:
            # For dimension-based filters, use actual width and height if provided
            if width_inches is not None and height_inches is not None:
                # Ensure width >= height (database convention)
                filter_width = max(width_inches, height_inches)
                filter_height = min(width_inches, height_inches)
                price = self._get_base_price_for_default_table(matching_filter, filter_width, filter_height)
            else:
                # Fallback: use size_inches for both dimensions (square filter)
                price = self._get_base_price_for_default_table(matching_filter, size_inches, size_inches)
            return price if price and price > 0 else None
    
    def _get_base_price_for_other_table(self, product, diameter_inches):
        """Get base price for other table product directly from database"""
        conn = self.price_loader._get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        cursor.execute('SELECT table_id FROM products WHERE model = ? LIMIT 1', (product,))
        table_result = cursor.fetchone()
        if not table_result:
            return None
        
        table_id = table_result[0]
        diameter_int = int(diameter_inches)
        
        cursor.execute('''
            SELECT width, normal_price
            FROM prices
            WHERE table_id = ? AND height IS NULL AND width >= ?
            ORDER BY width
            LIMIT 1
        ''', (table_id, diameter_int))
        
        result = cursor.fetchone()
        return result[1] if result else None
    
    def _get_base_price_for_default_table(self, product, width_inches, height_inches):
        """Get base price for default table product directly from database"""
        conn = self.price_loader._get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        cursor.execute('SELECT table_id FROM products WHERE model = ? LIMIT 1', (product,))
        table_result = cursor.fetchone()
        if not table_result:
            return None
        
        table_id = table_result[0]
        width_int = int(width_inches)
        height_int = int(height_inches)
        
        cursor.execute('''
            SELECT height, width, normal_price,
                   CASE 
                       WHEN height = ? AND width = ? THEN 0
                       ELSE (height - ?) + (width - ?)
                   END as priority
            FROM prices
            WHERE table_id = ? AND height >= ? AND width >= ?
            ORDER BY priority, height, width
            LIMIT 1
        ''', (height_int, width_int, height_int, width_int, table_id, height_int, width_int))
        
        result = cursor.fetchone()
        return result[2] if result else None

