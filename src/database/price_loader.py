"""
Price List Loader Module
Handles loading and parsing the SQLite price database for HRG and WSG products.
"""

import re
import sqlite3
import math
from pathlib import Path
from .equation_parser import EquationParser


class PriceListLoader:
    """Loads and parses the price list from SQLite database"""
    
    def __init__(self, db_path='../prices.db'):
        self.db_path = db_path
        self.conn = None
        self.equation_parser = EquationParser()
        self._check_database()
    
    def _is_number(self, value):
        """Check if a value is a simple number (not an equation)"""
        if value is None:
            return False
        try:
            float(str(value).strip())
            return True
        except ValueError:
            return False
    
    def _is_equation(self, value):
        """Check if a value is an equation (contains variables or functions)"""
        if value is None:
            return False
        value_str = str(value).strip()
        # Check if it contains equation-like patterns
        equation_indicators = ['TB', 'WD', 'WIDTH', 'HEIGHT', 'SIZE', '(', ')', '+', '-', '*', '/', 'sqrt', 'max', 'min', 'round', 'abs', 'ceil', 'floor', 'pow']
        return any(indicator in value_str for indicator in equation_indicators)
    
    def _calculate_hand_gear_addition(self, width, height, base_price):
        """
        Calculate hand gear addition for Model VD products.
        
        Args:
            width: Width dimension in inches
            height: Height dimension in inches
            base_price: Base table price from price list
            
        Returns:
            Total price including base price plus hand gear addition
        """
        import math
        # Hand gear calculation: (width*25/1000 rounded up) * (height*25/1000 rounded up) * 600
        width_factor = math.ceil((width * 25) / 1000)
        height_factor = math.ceil((height * 25) / 1000)
        hand_gear_addition = width_factor * height_factor * 600
        total_price = base_price + hand_gear_addition
        return total_price
    

    def _apply_tb_modifier(self, tb_modifier, tb_price, wd_price, variables):
        """Apply TB modifier to calculate base price (BP)"""
        bp_price = tb_price  # Default BP is same as TB
        if tb_modifier is not None:
            try:
                # Calculate base price using TB modifier equation
                if self._is_equation(tb_modifier):
                    bp_price = self.equation_parser.parse_equation(str(tb_modifier), variables)
                elif self._is_number(tb_modifier):
                    bp_price = tb_price * float(tb_modifier)
                else:
                    bp_price = tb_price
                
            except Exception as e:
                print(f"⚠ Warning: Error applying TB modifier: {e}")
                # Keep original bp_price if TB modifier fails
        
        return bp_price
    
    def _apply_wd_multiplier(self, wd_multiplier, wd_price, variables):
        """Apply WD multiplier to calculate modified WD price (MWD)"""
        modified_wd_price = wd_price  # Default modified WD is same as original WD
        if wd_multiplier is not None:
            try:
                # Calculate modified WD price using WD multiplier equation
                if self._is_equation(wd_multiplier):
                    modified_wd_price = self.equation_parser.parse_equation(str(wd_multiplier), variables)
                elif self._is_number(wd_multiplier):
                    modified_wd_price = wd_price * float(wd_multiplier)
                else:
                    modified_wd_price = wd_price
                
            except Exception as e:
                print(f"⚠ Warning: Error applying WD multiplier: {e}")
                # Keep original modified_wd_price if WD multiplier fails
        
        return modified_wd_price

    def _apply_pricing_logic(self, finish_multiplier, damper_multiplier, variables, with_damper=False):
        """Apply pricing logic: equations, multipliers, or fallback to base prices"""
        
        # Determine the base price to use
        if with_damper:
            base_price = variables['MWD']  # Use modified WD price for with-damper
        else:
            base_price = variables['BP']   # Use base price for without-damper
        
        # If there's a damper multiplier, apply it first
        if with_damper and damper_multiplier is not None:
            if self._is_equation(damper_multiplier):
                try:
                    base_price = self.equation_parser.parse_equation(str(damper_multiplier), variables)
                except Exception as e:
                    print(f"⚠ Warning: Error evaluating damper equation '{damper_multiplier}': {e}")
                    # Keep original base_price if equation fails
            elif self._is_number(damper_multiplier):
                base_price = base_price * float(damper_multiplier)
        
        # Apply finish multiplier if available
        if finish_multiplier is not None:
            if self._is_equation(finish_multiplier):
                try:
                    # Update variables with current base_price for equation evaluation
                    equation_variables = variables.copy()
                    equation_variables['BP'] = base_price
                    equation_variables['MWD'] = base_price
                    final_price = self.equation_parser.parse_equation(str(finish_multiplier), equation_variables)
                    return int(final_price + 0.5)
                except Exception as e:
                    print(f"⚠ Warning: Error evaluating finish equation '{finish_multiplier}': {e}")
                    return int(base_price + 0.5)
            elif self._is_number(finish_multiplier):
                final_price = base_price * float(finish_multiplier)
                return int(final_price + 0.5)
        
        # No finish multiplier, return base price
        return int(base_price + 0.5)

    
    def _get_exceeded_dimension_multiplier(self, conn, table_id, width, height, with_damper=False):
        """Get the appropriate multiplier when dimensions exceed table limits"""
        cursor = conn.cursor()
        
        # Get the maximum dimensions available in the table
        cursor.execute('''
            SELECT MAX(width), MAX(height)
            FROM prices
            WHERE table_id = ?
        ''', (table_id,))
        
        max_dims = cursor.fetchone()
        if not max_dims or not max_dims[0] or not max_dims[1]:
            return None
        
        max_width, max_height = max_dims
        
        # Check if both dimensions exceed limits - use fallback pricing
        if width > max_width and height > max_height:
            # Use the highest available multipliers as fallback
            # Try to get the highest row multiplier (for width exceeded)
            cursor.execute('''
                SELECT width_exceeded_multiplier, width_exceeded_multiplier_wd
                FROM row_multipliers
                WHERE table_id = ?
                ORDER BY height DESC
                LIMIT 1
            ''', (table_id,))
            row_result = cursor.fetchone()
            
            # Try to get the highest column multiplier (for height exceeded)
            cursor.execute('''
                SELECT height_exceeded_multiplier, height_exceeded_multiplier_wd
                FROM column_multipliers
                WHERE table_id = ?
                ORDER BY width DESC
                LIMIT 1
            ''', (table_id,))
            col_result = cursor.fetchone()
            
            # Use the higher of the two multipliers as fallback
            if row_result and col_result:
                if with_damper:
                    row_mult = row_result[1] if row_result[1] is not None else 0
                    col_mult = col_result[1] if col_result[1] is not None else 0
                else:
                    row_mult = row_result[0] if row_result[0] is not None else 0
                    col_mult = col_result[0] if col_result[0] is not None else 0
                
                # Use the higher multiplier
                return max(row_mult, col_mult) if max(row_mult, col_mult) > 0 else None
            
            return None
        
        # Check if only width exceeds limit - use specific height row multiplier
        elif width > max_width:
            # Find the closest height row that has a multiplier
            cursor.execute('''
                SELECT width_exceeded_multiplier, width_exceeded_multiplier_wd
                FROM row_multipliers
                WHERE table_id = ? AND height <= ?
                ORDER BY height DESC
                LIMIT 1
            ''', (table_id, height))
            result = cursor.fetchone()
            if result:
                # Return appropriate multiplier based on with_damper flag
                if with_damper and result[1] is not None:
                    return result[1]  # WD multiplier
                elif not with_damper and result[0] is not None:
                    return result[0]  # Regular multiplier
            return None
        
        # Check if only height exceeds limit - use specific width column multiplier
        elif height > max_height:
            # Find the closest width column that has a multiplier
            cursor.execute('''
                SELECT height_exceeded_multiplier, height_exceeded_multiplier_wd
                FROM column_multipliers
                WHERE table_id = ? AND width <= ?
                ORDER BY width DESC
                LIMIT 1
            ''', (table_id, width))
            result = cursor.fetchone()
            if result:
                # Return appropriate multiplier based on with_damper flag
                if with_damper and result[1] is not None:
                    return result[1]  # WD multiplier
                elif not with_damper and result[0] is not None:
                    return result[0]  # Regular multiplier
            return None
        
        # No dimensions exceeded
        return None
    
    
    def _check_database(self):
        """Check if database exists and is accessible"""
        if not Path(self.db_path).exists():
            print(f"❌ Error: Database file '{self.db_path}' not found!")
            return False
        return True
    
    def _get_connection(self):
        """Get database connection, creating one if needed"""
        if self.conn is None:
            if not self._check_database():
                return None
            self.conn = sqlite3.connect(self.db_path)
        return self.conn
    
    def get_available_models(self):
        """Get list of available product models"""
        conn = self._get_connection()
        if not conn:
            return []
        
        cursor = conn.cursor()
        cursor.execute('SELECT DISTINCT model FROM products ORDER BY model')
        return [row[0] for row in cursor.fetchall()]
    
    def get_available_finishes(self, product):
        """Get list of available finish options for a specific product"""
        conn = self._get_connection()
        if not conn:
            return []
        
        cursor = conn.cursor()
        cursor.execute('''
            SELECT anodized_multiplier, powder_coated_multiplier 
            FROM products 
            WHERE model = ? 
            LIMIT 1
        ''', (product,))
        
        result = cursor.fetchone()
        if not result:
            return []
        
        available_finishes = []
        anodized_multiplier, powder_coated_multiplier = result
        
        # Check if anodized aluminum is available (multiplier is not None)
        if anodized_multiplier is not None:
            available_finishes.append('Anodized Aluminum')
        
        # Check if powder coated is available (multiplier is not None)
        if powder_coated_multiplier is not None:
            available_finishes.append('Powder Coated')
        
        # Always add Special Color option
        available_finishes.append('Special Color')
        
        return available_finishes
    
    def is_other_table(self, product):
        """Check if a product uses other table format (diameter-based) instead of width/height"""
        conn = self._get_connection()
        if not conn:
            return False
        
        cursor = conn.cursor()
        cursor.execute('''
            SELECT COUNT(*) 
            FROM products p
            JOIN prices pr ON p.table_id = pr.table_id
            WHERE p.model = ? AND pr.width IS NULL
        ''', (product,))
        
        result = cursor.fetchone()
        return result[0] > 0 if result else False
    
    def get_price_for_default_table(self, product, finish, size, with_damper=False, special_color_multiplier=1.0):
        """Get price for a specific product configuration using idx_price_lookup index"""
        conn = self._get_connection()
        if not conn:
            return 0
        
        # Parse size to get width and height
        size_match = re.search(r'(\d+(?:\.\d+)?)"?\s*x\s*(\d+(?:\.\d+)?)"?', size.lower())
        if not size_match:
            return 0
        
        width = int(float(size_match.group(1)))
        height = int(float(size_match.group(2)))
        
        cursor = conn.cursor()
        
        # Get table_id, multipliers, and equations for the product
        cursor.execute('SELECT table_id, tb_modifier, anodized_multiplier, powder_coated_multiplier, wd_multiplier FROM products WHERE model = ? LIMIT 1', (product,))
        table_result = cursor.fetchone()
        if not table_result:
            return 0
        
        table_id = table_result[0]
        tb_modifier = table_result[1]
        anodized_multiplier = table_result[2]
        powder_coated_multiplier = table_result[3]
        wd_multiplier = table_result[4]
        
        # Special case for Model VD: Check for oversized dimensions first
        if product == 'VD':
            # Check for VD oversized dimensions using inch values directly
            vd_oversized_price = self._calculate_vd_oversized_price(conn, table_id, width, height, with_damper)
            if vd_oversized_price is not None:
                # Use the VD oversized price directly
                return int(vd_oversized_price + 0.5)
        
        # Check if dimensions exceed table limits and get appropriate multipliers
        exceeded_tb_multiplier = self._get_exceeded_dimension_multiplier(conn, table_id, width, height, with_damper=False)
        exceeded_wd_multiplier = self._get_exceeded_dimension_multiplier(conn, table_id, width, height, with_damper=True)
        
        # Get base prices - either from exceeded dimension calculation or database lookup
        if exceeded_tb_multiplier is not None:
            # Calculate base prices using exceeded dimension formula: height * width * multiplier
            tb_price = height * width * exceeded_tb_multiplier
            wd_price = height * width * exceeded_wd_multiplier if exceeded_wd_multiplier is not None else tb_price
        else:
            # Query prices using the table_id and dimensions - always return normal_price first, then price_with_damper
            cursor.execute('''
                SELECT normal_price, price_with_damper
                FROM prices
                WHERE table_id = ? AND width = ? AND height = ?
            ''', (table_id, width, height))
            
            result = cursor.fetchone()
            if not result or (result[0] is None and result[1] is None):
                return 0
            
            # Get base prices
            tb_price = result[0] if result[0] is not None else 0  # Table base price
            wd_price = result[1] if result[1] is not None else 0   # With damper price
        
        # Create variables for calculations
        tb_variables = {
            'TB': tb_price,
            'WD': wd_price,
            'WIDTH': width,
            'HEIGHT': height
        }
        
        # Apply TB modifier to calculate base price (BP)
        bp_price = self._apply_tb_modifier(tb_modifier, tb_price, wd_price, tb_variables)
        
        # Create variables for WD multiplier calculation
        wd_variables = {
            'TB': tb_price,
            'WD': wd_price,
            'BP': bp_price,
            'WIDTH': width,
            'HEIGHT': height
        }
        
        # Apply WD multiplier to calculate modified WD price (MWD)
        modified_wd_price = self._apply_wd_multiplier(wd_multiplier, wd_price, wd_variables)
        
        # Determine which multiplier/equation to use based on finish and damper option
        finish_multiplier = None
        wd_multiplier_value = None
        
        # Get finish multiplier
        if finish == 'Anodized Aluminum' and anodized_multiplier is not None:
            finish_multiplier = anodized_multiplier
        elif 'Powder Coated' in finish and powder_coated_multiplier is not None:
            # Powder Coated uses the same multiplier regardless of color
            finish_multiplier = powder_coated_multiplier
        elif 'Special Color' in finish:
            # Use the user-provided multiplier for special colors
            finish_multiplier = special_color_multiplier
        
        # Get WD multiplier
        if wd_multiplier is not None:
            wd_multiplier_value = wd_multiplier
        
        # Calculate final price
        variables = {
            'TB': tb_price,           # TB remains the original table price
            'WD': wd_price,           # WD remains the original with-damper price
            'BP': bp_price,           # BP (Base Price) is the calculated value from TB modifier
            'MWD': modified_wd_price, # MWD (Modified WD) is the calculated value from WD multiplier
            'WIDTH': width,
            'HEIGHT': height
        }
        
        # Apply pricing logic with both multipliers
        final_price = self._apply_pricing_logic(finish_multiplier, wd_multiplier_value, variables, with_damper)
        
        # Special case for Model VD: Add hand gear calculation to final price
        if product == 'VD':
            final_price = self._calculate_hand_gear_addition(width, height, final_price)
        
        return final_price
    
    def find_rounded_default_table_size(self, product, finish, width, height):
        """Find the exact match first, then the next available size that is >= the given width and height"""
        conn = self._get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        
        # Single query: prioritize exact match, then closest >= match
        cursor.execute('''
            SELECT pr.width, pr.height,
                   CASE 
                       WHEN pr.width = ? AND pr.height = ? THEN 0
                       ELSE (pr.width - ?) + (pr.height - ?)
                   END as priority
            FROM products p
            JOIN prices pr ON p.table_id = pr.table_id
            WHERE p.model = ? AND pr.width >= ? AND pr.height >= ?
            ORDER BY priority, pr.width, pr.height
            LIMIT 1
        ''', (width, height, width, height, product, width, height))
        
        result = cursor.fetchone()
        if result:
            return f'{result[0]}" x {result[1]}"'
        
        return None
    
    def get_price_for_other_table(self, product, finish, diameter, with_damper=False, special_color_multiplier=1.0):
        """Get price for an other table (diameter-based) product configuration"""
        conn = self._get_connection()
        if not conn:
            return 0
        
        cursor = conn.cursor()
        
        # Parse diameter value if it's a string (e.g., "8\" diameter" -> 8)
        if isinstance(diameter, str):
            diameter_match = re.search(r'(\d+(?:\.\d+)?)', diameter)
            if not diameter_match:
                return 0
            diameter = int(float(diameter_match.group(1)))
        
        # Get table_id, multipliers, and equations for the product
        cursor.execute('SELECT table_id, tb_modifier, anodized_multiplier, powder_coated_multiplier, wd_multiplier FROM products WHERE model = ? LIMIT 1', (product,))
        table_result = cursor.fetchone()
        if not table_result:
            return 0
        
        table_id = table_result[0]
        tb_modifier = table_result[1]
        anodized_multiplier = table_result[2]
        powder_coated_multiplier = table_result[3]
        wd_multiplier = table_result[4]
        
        # Query prices using the table_id and diameter (height field for diameter-based products) - always return normal_price first, then price_with_damper
        cursor.execute('''
            SELECT normal_price, price_with_damper
            FROM prices
            WHERE table_id = ? AND width IS NULL AND height = ?
        ''', (table_id, diameter))
        
        result = cursor.fetchone()
        if not result or (result[0] is None and result[1] is None):
            return 0
        
        # Get base prices
        tb_price = result[0] if result[0] is not None else 0  # Table base price
        wd_price = result[1] if result[1] is not None else 0   # With damper price
        
        # Create variables for calculations
        tb_variables = {
            'TB': tb_price,
            'WD': wd_price,
            'SIZE': diameter
        }
        
        # Apply TB modifier to calculate base price (BP)
        bp_price = self._apply_tb_modifier(tb_modifier, tb_price, wd_price, tb_variables)
        
        # Create variables for WD multiplier calculation
        wd_variables = {
            'TB': tb_price,
            'WD': wd_price,
            'BP': bp_price,
            'SIZE': diameter
        }
        
        # Apply WD multiplier to calculate modified WD price (MWD)
        modified_wd_price = self._apply_wd_multiplier(wd_multiplier, wd_price, wd_variables)
        
        # Determine which multiplier/equation to use based on finish and damper option
        finish_multiplier = None
        wd_multiplier_value = None
        
        # Get finish multiplier
        if finish == 'Anodized Aluminum' and anodized_multiplier is not None:
            finish_multiplier = anodized_multiplier
        elif 'Powder Coated' in finish and powder_coated_multiplier is not None:
            # Powder Coated uses the same multiplier regardless of color
            finish_multiplier = powder_coated_multiplier
        elif 'Special Color' in finish:
            # Use the user-provided multiplier for special colors
            finish_multiplier = special_color_multiplier
        
        # Get WD multiplier
        if wd_multiplier is not None:
            wd_multiplier_value = wd_multiplier
        
        # Calculate final price
        variables = {
            'TB': tb_price,           # TB remains the original table price
            'WD': wd_price,           # WD remains the original with-damper price
            'BP': bp_price,           # BP (Base Price) is the calculated value from TB modifier
            'MWD': modified_wd_price, # MWD (Modified WD) is the calculated value from WD multiplier
            'SIZE': diameter
        }
        return self._apply_pricing_logic(finish_multiplier, wd_multiplier_value, variables, with_damper)
    
    def find_rounded_other_table_size(self, product, finish, diameter):
        """Find the exact match first, then the next available diameter that is >= the given diameter for other table products"""
        conn = self._get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        
        # Single query: prioritize exact match, then closest >= match
        cursor.execute('''
            SELECT pr.height,
                   CASE 
                       WHEN pr.height = ? THEN 0
                       ELSE pr.height - ?
                   END as priority
            FROM products p
            JOIN prices pr ON p.table_id = pr.table_id
            WHERE p.model = ? AND pr.width IS NULL AND pr.height >= ?
            ORDER BY priority, pr.height
            LIMIT 1
        ''', (diameter, diameter, product, diameter))
        
        result = cursor.fetchone()
        if result:
            return f'{result[0]}" diameter'
        
        return None
    
    def has_damper_option(self, product, finish):
        """Check if a product has a non-null WD multiplier in the header sheet"""
        conn = self._get_connection()
        if not conn:
            return False
        
        cursor = conn.cursor()
        
        # Get WD multiplier for the product from the products table
        cursor.execute('SELECT wd_multiplier FROM products WHERE model = ? LIMIT 1', (product,))
        result = cursor.fetchone()
        if not result:
            return False
        
        wd_multiplier = result[0]
        # Return True only if WD multiplier is not None and not empty
        return wd_multiplier is not None and str(wd_multiplier).strip() != ''
    
    def _calculate_vd_oversized_price(self, conn, table_id, width, height, with_damper=False):
        """
        Calculate price for VD products when dimensions exceed limits:
        - Height > 80 inch or Width > 40 inch
        - Use maximum dimensions (80 inch height, 40 inch width) to calculate round-up multipliers
        - Find price for the divided dimensions and multiply by round-up numbers
        """
        # Check if dimensions exceed VD limits (using inch values directly)
        height_exceeds = height > 80  # 80 inch
        width_exceeds = width > 40   # 40 inch
        
        if not height_exceeds and not width_exceeds:
            return None  # No oversized calculation needed
        
        # Calculate round-up multipliers
        height_multiplier = 1
        width_multiplier = 1
        
        if height_exceeds:
            height_multiplier = math.ceil(height / 80)
            adjusted_height = height / height_multiplier
        else:
            adjusted_height = height
        
        if width_exceeds:
            width_multiplier = math.ceil(width / 40)
            adjusted_width = width / width_multiplier
        else:
            adjusted_width = width
        
        # Convert adjusted dimensions to integers for database lookup
        adjusted_height_inches = int(adjusted_height)
        adjusted_width_inches = int(adjusted_width)
        
        # Find the closest available dimensions in the database
        cursor = conn.cursor()
        
        # First, get the maximum available dimensions
        cursor.execute('''
            SELECT MAX(width), MAX(height)
            FROM prices
            WHERE table_id = ?
        ''', (table_id,))
        
        max_dims = cursor.fetchone()
        if not max_dims or not max_dims[0] or not max_dims[1]:
            return None
        
        max_width, max_height = max_dims
        
        # Use the maximum available dimensions if adjusted dimensions exceed them
        lookup_width = min(adjusted_width_inches, max_width)
        lookup_height = min(adjusted_height_inches, max_height)
        
        # Find the price for the lookup dimensions
        cursor.execute('''
            SELECT normal_price, price_with_damper
            FROM prices
            WHERE table_id = ? AND width = ? AND height = ?
        ''', (table_id, lookup_width, lookup_height))
        
        result = cursor.fetchone()
        
        # If no exact match found, try to find the closest available dimensions
        if not result or (result[0] is None and result[1] is None):
            # Try to find the closest available dimensions
            cursor.execute('''
                SELECT width, height, normal_price, price_with_damper
                FROM prices
                WHERE table_id = ? AND width <= ? AND height <= ?
                ORDER BY width DESC, height DESC
                LIMIT 1
            ''', (table_id, lookup_width, lookup_height))
            
            result = cursor.fetchone()
            
            if result:
                lookup_width = result[0]
                lookup_height = result[1]
                # Extract prices from the result
                normal_price = result[2]
                price_with_damper = result[3]
                result = (normal_price, price_with_damper)
        
        if not result or (result[0] is None and result[1] is None):
            return None
        
        # Get the appropriate price based on with_damper flag
        if with_damper and result[1] is not None:
            base_price = result[1]
        elif not with_damper and result[0] is not None:
            base_price = result[0]
        else:
            # Fallback to available price
            base_price = result[0] if result[0] is not None else result[1]
        
        if base_price is None:
            return None
        
        # Calculate total multiplier
        total_multiplier = height_multiplier * width_multiplier
        
        # Calculate final price
        final_price = base_price * total_multiplier
        
        # Add hand gear calculation for VD products
        hand_gear_price = self._calculate_hand_gear_addition(width, height, final_price)
        
        return hand_gear_price
    
    def __del__(self):
        """Clean up database connection when object is destroyed"""
        if self.conn:
            self.conn.close()
