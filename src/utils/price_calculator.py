"""
Price Calculator Module
Handles price calculations using data from the SQLite price database for HRG and WSG products.
"""

import re
import math
from utils.equation_parser import EquationParser
from utils.sql_loader import PriceDatabase


class PriceCalculator:
    """Calculates prices using data from SQLite database"""
    
    def __init__(self, db_path='../prices.db'):
        self.db_path = db_path
        self.db = PriceDatabase(db_path)
        self.equation_parser = EquationParser()
    
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
        # Hand gear calculation: rounded up(height/1500) x rounded up(width/1500) x hand gear price
        hand_gear_price = 600
        height_factor = math.ceil((height * 25) / 1500)
        width_factor = math.ceil((width * 25) / 1500)
        hand_gear_addition = height_factor * width_factor * hand_gear_price
        return base_price + hand_gear_addition
    

    def _apply_modifier(self, modifier, base_value, variables, default_value=None):
        """
        Apply a modifier (equation or multiplier) to a base value.
        
        Args:
            modifier: Modifier value (can be equation string, number, or None)
            base_value: Base value to apply modifier to
            variables: Variables dictionary for equation evaluation
            default_value: Default value if modifier is None (defaults to base_value)
            
        Returns:
            Modified value
        """
        if default_value is None:
            default_value = base_value
        
        if modifier is None:
            return default_value
        
        try:
            if self._is_equation(modifier):
                return self.equation_parser.parse_equation(str(modifier), variables)
            elif self._is_number(modifier):
                return base_value * float(modifier)
            else:
                return default_value
        except Exception as e:
            print(f"⚠ Warning: Error applying modifier: {e}")
            return default_value
    
    def _apply_tb_modifier(self, tb_modifier, tb_price, wd_price, variables):
        """Apply TB modifier to calculate base price (BP)"""
        return self._apply_modifier(tb_modifier, tb_price, variables, tb_price)
    
    def _apply_wd_multiplier(self, wd_multiplier, wd_price, variables):
        """Apply WD multiplier to calculate modified WD price (MWD)"""
        return self._apply_modifier(wd_multiplier, wd_price, variables, wd_price)

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

    
    def _get_exceeded_dimension_multiplier(self, table_id, width, height, with_damper=False):
        """Get the appropriate multiplier when dimensions exceed table limits"""
        return self.db.get_exceeded_dimension_multiplier(table_id, width, height, with_damper)
    
    def get_available_models(self):
        """Get list of available product models"""
        return self.db.get_available_models()
    
    def get_available_finishes(self, product):
        """Get list of available finish options for a specific product"""
        return self.db.get_available_finishes(product)
    
    def is_other_table(self, product):
        """Check if a product uses other table format (diameter-based) instead of width/height"""
        return self.db.is_other_table(product)
    
    def has_price_per_foot(self, product):
        """Check if a product has price_per_foot pricing"""
        return self.db.has_price_per_foot(product)
    
    def has_no_dimensions(self, product):
        """Check if a product has no height and width (both are NULL)"""
        return self.db.has_no_dimensions(product)
    
    def get_price_id_for_no_dimensions(self, product):
        """Get price_id for a product with no height/width by calculating from product_id difference
        
        Args:
            product: Product model name
            
        Returns:
            price_id or None if not found
        """
        return self.db.get_price_id_for_no_dimensions(product)
    
    def get_price_for_price_per_foot(self, product, finish, width, height, with_damper=False, special_color_multiplier=1.0, price_id=None):
        """Get price for a product with price_per_foot pricing
        
        Formula: (height / 12) × price_per_foot (matching width first)
        
        Args:
            product: Product model name
            finish: Finish type
            width: Width in inches (must match database width) - not used if price_id is provided
            height: Height in inches
            with_damper: Whether product has damper option
            special_color_multiplier: Multiplier for special color finish
            price_id: Optional price_id to use directly (for has_no_dimensions case)
            
        Returns:
            Calculated price or 0 if not found
        """
        # Get product multipliers
        multipliers = self.db.get_product_multipliers(product)
        if not multipliers:
            return 0
        
        anodized_multiplier = multipliers[0]
        powder_coated_multiplier = multipliers[1]
        
        # Get table_id
        table_id = self.db.get_table_id(product)
        if table_id is None:
            return 0
        
        # Get price_per_foot
        # Note: 'width' parameter is actually the size value stored in SQL height column
        if price_id is not None:
            price_per_foot = self.db.get_price_per_foot(table_id, 0, price_id)
        else:
            size_int = int(width)  # This is the size from SQL height column
            price_per_foot = self.db.get_price_per_foot(table_id, size_int)
        
        if price_per_foot is None:
            return 0
        
        # Calculate base price: (height / 12) * price_per_foot
        # Note: 'height' parameter is the dimension to multiply with price_per_foot
        base_price = (height / 12) * price_per_foot
        
        # Apply finish multiplier
        finish_multiplier = 1.0
        if finish == 'Anodized Aluminum' and anodized_multiplier is not None:
            if self._is_equation(anodized_multiplier):
                variables = {
                    'BP': base_price,
                    'WIDTH': width,
                    'HEIGHT': height
                }
                try:
                    finish_multiplier = self.equation_parser.parse_equation(str(anodized_multiplier), variables)
                except Exception as e:
                    print(f"⚠ Warning: Error evaluating anodized equation: {e}")
                    finish_multiplier = 1.0
            elif self._is_number(anodized_multiplier):
                finish_multiplier = float(anodized_multiplier)
        elif 'Powder Coated' in finish and powder_coated_multiplier is not None:
            if self._is_equation(powder_coated_multiplier):
                variables = {
                    'BP': base_price,
                    'WIDTH': width,
                    'HEIGHT': height
                }
                try:
                    finish_multiplier = self.equation_parser.parse_equation(str(powder_coated_multiplier), variables)
                except Exception as e:
                    print(f"⚠ Warning: Error evaluating powder coated equation: {e}")
                    finish_multiplier = 1.0
            elif self._is_number(powder_coated_multiplier):
                finish_multiplier = float(powder_coated_multiplier)
        elif 'Special Color' in finish:
            finish_multiplier = special_color_multiplier
        
        final_price = base_price * finish_multiplier
        return int(final_price + 0.5)
    
    def find_rounded_price_per_foot_width(self, product, width):
        """Find the exact match first, then the next available width that is >= the given width for price_per_foot products"""
        width_int = int(width)
        return self.db.find_rounded_price_per_foot_width(product, width_int)
    
    def get_price_for_default_table(self, product, finish, size, with_damper=False, special_color_multiplier=1.0):
        """Get price for a specific product configuration using idx_price_lookup index"""
        # Parse size to get width and height
        size_match = re.search(r'(\d+(?:\.\d+)?)"?\s*x\s*(\d+(?:\.\d+)?)"?', size.lower())
        if not size_match:
            return 0
        
        width = int(float(size_match.group(1)))
        height = int(float(size_match.group(2)))
        
        # Get product data
        product_data = self.db.get_product_data(product)
        if not product_data:
            return 0
        
        table_id = product_data[0]
        tb_modifier = product_data[1]
        anodized_multiplier = product_data[2]
        powder_coated_multiplier = product_data[3]
        wd_multiplier = product_data[4]
        
        # Special case for Model VD: Check for oversized dimensions first
        if product == 'VD':
            # Check for VD oversized dimensions using inch values directly
            vd_oversized_price = self._calculate_vd_oversized_price(table_id, width, height, with_damper)
            if vd_oversized_price is not None:
                # Use the VD oversized price directly
                return int(vd_oversized_price + 0.5)
        
        # Check if dimensions exceed table limits and get appropriate multipliers
        exceeded_tb_multiplier = self._get_exceeded_dimension_multiplier(table_id, width, height, with_damper=False)
        exceeded_wd_multiplier = self._get_exceeded_dimension_multiplier(table_id, width, height, with_damper=True)
        
        # Get base prices - either from exceeded dimension calculation or database lookup
        if exceeded_tb_multiplier is not None:
            # Calculate base prices using exceeded dimension formula: height * width * multiplier
            tb_price = height * width * exceeded_tb_multiplier
            wd_price = height * width * exceeded_wd_multiplier if exceeded_wd_multiplier is not None else tb_price
        else:
            # Query prices using the table_id and dimensions
            price_result = self.db.get_price_for_dimensions(table_id, height, width)
            if not price_result:
                return 0
            
            # Get base prices
            tb_price = price_result[0] if price_result[0] is not None else 0  # Table base price
            wd_price = price_result[1] if price_result[1] is not None else 0   # With damper price
        
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
        return self.db.find_rounded_default_table_size(product, width, height)
    
    def get_price_for_other_table(self, product, finish, diameter, with_damper=False, special_color_multiplier=1.0):
        """Get price for an other table (diameter-based) product configuration"""
        # Parse diameter value if it's a string (e.g., "8\" diameter" -> 8)
        if isinstance(diameter, str):
            diameter_match = re.search(r'(\d+(?:\.\d+)?)', diameter)
            if not diameter_match:
                return 0
            diameter = int(float(diameter_match.group(1)))
        
        # Get product data
        product_data = self.db.get_product_data(product)
        if not product_data:
            return 0
        
        table_id = product_data[0]
        tb_modifier = product_data[1]
        anodized_multiplier = product_data[2]
        powder_coated_multiplier = product_data[3]
        wd_multiplier = product_data[4]
        
        # Query prices using the table_id and diameter
        price_result = self.db.get_price_for_diameter(table_id, diameter)
        if not price_result:
            return 0
        
        # Get base prices
        tb_price = price_result[0] if price_result[0] is not None else 0  # Table base price
        wd_price = price_result[1] if price_result[1] is not None else 0   # With damper price
        
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
    
    def find_rounded_other_table_size(self, product, finish, diameter, price_id=None):
        """Find the exact match first, then the next available diameter that is >= the given diameter for other table products
        
        Args:
            product: Product model name
            finish: Finish type (not used, kept for compatibility)
            diameter: Diameter in inches (not used if price_id is provided)
            price_id: Optional price_id to use directly (for has_no_dimensions case)
            
        Returns:
            Size string (e.g., "8\" diameter") or None if not found
        """
        diameter_int = int(diameter) if not isinstance(diameter, str) else int(float(re.search(r'(\d+(?:\.\d+)?)', diameter).group(1)))
        return self.db.find_rounded_other_table_size(product, diameter_int, price_id)
    
    def has_damper_option(self, product, finish):
        """Check if a product has a non-null WD multiplier in the header sheet"""
        return self.db.has_damper_option(product)
    
    def _calculate_vd_oversized_price(self, table_id, width, height, with_damper=False):
        """
        Calculate price for VD products when dimensions exceed limits:
        - Height > 40 inch or Width > 80 inch
        - Use maximum dimensions (40 inch height, 80 inch width) to calculate round-up multipliers
        - Find price for the divided dimensions and multiply by round-up numbers
        """
        # Check if dimensions exceed VD limits (using inch values directly)
        height_exceeds = height > 40  # 40 inch
        width_exceeds = width > 80   # 80 inch
        
        if not height_exceeds and not width_exceeds:
            return None  # No oversized calculation needed
        
        # Calculate round-up multipliers
        height_multiplier = 1
        width_multiplier = 1
        
        if height_exceeds:
            height_multiplier = math.ceil(height / 40)
            adjusted_height = height / height_multiplier
        else:
            adjusted_height = height
        
        if width_exceeds:
            width_multiplier = math.ceil(width / 80)
            adjusted_width = width / width_multiplier
        else:
            adjusted_width = width
        
        # Convert adjusted dimensions to integers for database lookup
        adjusted_height_inches = int(adjusted_height)
        adjusted_width_inches = int(adjusted_width)
        
        # Get maximum available dimensions
        max_dims = self.db.get_max_dimensions(table_id)
        if not max_dims:
            return None
        
        max_height, max_width = max_dims
        
        # Use the maximum available dimensions if adjusted dimensions exceed them
        lookup_width = min(adjusted_width_inches, max_width)
        lookup_height = min(adjusted_height_inches, max_height)
        
        # Find the price for the lookup dimensions
        price_result = self.db.get_price_for_dimensions(table_id, lookup_height, lookup_width)
        
        # If no exact match found, try to find the closest available dimensions
        if not price_result:
            closest = self.db.find_closest_price_for_dimensions(table_id, lookup_height, lookup_width)
            if closest:
                lookup_height = closest[0]
                lookup_width = closest[1]
                price_result = (closest[2], closest[3])
        
        if not price_result or (price_result[0] is None and price_result[1] is None):
            return None
        
        # Get the appropriate price based on with_damper flag
        if with_damper and price_result[1] is not None:
            base_price = price_result[1]
        elif not with_damper and price_result[0] is not None:
            base_price = price_result[0]
        else:
            # Fallback to available price
            base_price = price_result[0] if price_result[0] is not None else price_result[1]
        
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
        if hasattr(self, 'db'):
            self.db.close()
