"""
Price Calculator Module
Handles price calculations using data from the SQLite price database for HRG and WSG products.
"""

import re
import math
from utils.equation_parser import EquationParser
from utils.sql_loader import PriceDatabase


class PriceNotFoundError(Exception):
    """Raised when a price cannot be found in the database"""
    pass


class ProductNotFoundError(Exception):
    """Raised when a product cannot be found in the database"""
    pass


class SizeNotFoundError(Exception):
    """Raised when a size cannot be found in the database"""
    pass


class ModifierError(Exception):
    """Raised when there's an error applying a modifier"""
    pass


class PriceCalculator:
    """Calculates prices using data from SQLite database"""
    
    def __init__(self, db_path='../prices.db'):
        self.db_path = db_path
        self.db = PriceDatabase(db_path)
        self.equation_parser = EquationParser()
    
    def get_hand_gear_price(self, product, width, height):
        """
        Get hand gear price for VD, VD-G, and VD-M products.
        
        Args:
            product: Product model name (VD, VD-G, or VD-M)
            width: Width dimension in inches
            height: Height dimension in inches
            
        Returns:
            Hand gear price amount (0 if product doesn't use hand gear)
        """
        if product == 'VD-G' or product == 'RVD-G':
            hand_gear_price = 600
        elif product == 'VD-M' or product == 'RVD-M':
            hand_gear_price = 7000
        else:
            return 0
        
        # Hand gear calculation: rounded up(height/1500) x rounded up(width/1500) x hand gear price
        # Convert inches to mm using 25.4 factor
        height_factor = math.ceil((height * 25.4) / 1500)
        width_factor = math.ceil((width * 25.4) / 1500)
        hand_gear_addition = height_factor * width_factor * hand_gear_price
        return hand_gear_addition
    

    def _apply_modifier(self, modifier, base_value, variables):
        """
        Apply a modifier (equation or multiplier) to a base value.
        
        Args:
            modifier: Modifier value (can be equation string, number, or None)
            base_value: Base value to apply modifier to
            variables: Variables dictionary for equation evaluation
            
        Returns:
            Modified value
        """
        if modifier is None:
            return base_value
        
        try:
            if self.equation_parser.is_equation(modifier):
                return self.equation_parser.parse_equation(str(modifier), variables)
            elif self.equation_parser.is_number(modifier):
                return base_value * float(modifier)
            else:
                return base_value
        except Exception as e:
            raise ModifierError(f"Error applying modifier '{modifier}': {e}") from e
    
    def _apply_base_modifier(self, base_modifier, variables):
        """Apply Base modifier to calculate base price (BP)"""
        tb_price = variables.get('TB', 0)
        return self._apply_modifier(base_modifier, tb_price, variables)
    
    def _apply_wd_modifier(self, wd_modifier, variables):
        """Apply WD modifier to calculate modified WD price (MWD)"""
        wd_price = variables.get('WD', 0)
        return self._apply_modifier(wd_modifier, wd_price, variables)

    def _apply_finish_pricing(self, finish_multiplier, variables, with_damper=False, finish=None):
        """Apply finish multiplier to base price (BP or MWD) and return final price"""
        if with_damper:
            base_price = variables.get('MWD', 0)  # Use modified WD price for with-damper
        else:
            base_price = variables.get('BP', 0)   # Use base price for without-damper
        
        return self._apply_modifier(finish_multiplier, base_price, variables)
    
    def _load_product_data(self, product):
        """
        Load and extract product data from the database.
        
        Args:
            product: Product model name
            
        Returns:
            Tuple of (table_id, base_modifier, anodized_multiplier, powder_coated_multiplier, no_finish_multiplier, wd_modifier)
            
        Raises:
            ProductNotFoundError: If product data not found
        """
        product_data = self.db.get_product_data(product)
        if not product_data:
            raise ProductNotFoundError(f'Product data not found for {product}')
        
        table_id = product_data[0]
        base_modifier = product_data[1]
        anodized_multiplier = product_data[2]
        powder_coated_multiplier = product_data[3]
        no_finish_multiplier = product_data[4]
        wd_modifier = product_data[5]
        
        return table_id, base_modifier, anodized_multiplier, powder_coated_multiplier, no_finish_multiplier, wd_modifier
    
    def _calculate_final_price_from_base_prices(self, tb_price, wd_price, base_modifier, anodized_multiplier, 
                                                 powder_coated_multiplier, no_finish_multiplier, wd_modifier, finish, with_damper=False, 
                                                 special_color_multiplier=None):
        """
        Calculate final price from base TB and WD prices using modifiers and finish multipliers.
        This is a shared helper function used by get_price_for_other_table, get_price_for_default_table,
        and get_price_for_price_per_foot.
        
        Args:
            tb_price: Table base price (TB)
            wd_price: With damper price (WD)
            base_modifier: Base modifier from product data
            anodized_multiplier: Anodized multiplier from product data
            powder_coated_multiplier: Powder coated multiplier from product data
            no_finish_multiplier: No finish multiplier from product data
            wd_modifier: WD modifier from product data
            finish: Finish type
            with_damper: Whether product has damper option
            special_color_multiplier: Multiplier for special color finish
            
        Returns:
            Final calculated price
        """
        # Create variables for Base modifier calculation
        tb_variables = {
            'TB': tb_price,
            'WD': wd_price
        }
        
        # Apply Base modifier to calculate base price (BP)
        bp_price = self._apply_base_modifier(base_modifier, tb_variables)
        
        # Create variables for WD multiplier calculation
        wd_variables = {
            'TB': tb_price,
            'WD': wd_price,
            'BP': bp_price
        }
        
        # Apply WD modifier to calculate modified WD price (MWD) only when with_damper is True
        if with_damper:
            modified_wd_price = self._apply_wd_modifier(wd_modifier, wd_variables)
        else:
            modified_wd_price = 0
        
        # Determine which multiplier/equation to use based on finish and damper option
        finish_multiplier = None
        
        # Get finish multiplier
        # Handle None finish (no finish applied, multiplier stays 1.0)
        if finish is None:
            finish_multiplier = None  # Will result in multiplier 1.0
        elif finish and 'No Finish' in finish and no_finish_multiplier is not None:
            # No Finish uses the same multiplier regardless of finish_str
            finish_multiplier = no_finish_multiplier
        elif finish and 'Anodized Aluminum' in finish and anodized_multiplier is not None:
            # Anodized Aluminum uses the same multiplier regardless of finish_str
            finish_multiplier = anodized_multiplier
        elif finish and 'Powder Coated' in finish and powder_coated_multiplier is not None:
            # Powder Coated uses the same multiplier regardless of color
            finish_multiplier = powder_coated_multiplier
        elif finish and 'Special Color' in finish:
            # Use the user-provided multiplier for special colors
            if special_color_multiplier is None:
                finish_multiplier = 1.45
            else:
                finish_multiplier = special_color_multiplier
        
        # Calculate final price
        variables = {
            'TB': tb_price,           # TB remains the original table price
            'WD': wd_price,           # WD remains the original with-damper price
            'BP': bp_price,           # BP (Base Price) is the calculated value from TB modifier
            'MWD': modified_wd_price, # MWD (Modified WD) is the calculated value from WD multiplier
        }
        
        # Apply finish pricing - MWD already has the WD modifier applied
        final_price = self._apply_finish_pricing(finish_multiplier, variables, with_damper)
        
        return final_price, finish_multiplier

    
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
    
    def has_price_per_sq_in(self, product):
        """Check if a product has price_per_sq_in pricing"""
        return self.db.has_price_per_sq_in(product)
    
    def has_no_dimensions(self, product):
        """Check if a product has no height and width (both are NULL)"""
        return self.db.has_no_dimensions(product)
    
    def get_price_id_for_no_dimensions(self, product):
        """Get price_id for a product with no height/width by calculating from product_id difference
        
        Args:
            product: Product model name
            
        Returns:
            price_id
            
        Raises:
            ProductNotFoundError: If price_id is not found
        """
        price_id = self.db.get_price_id_for_no_dimensions(product)
        if price_id is None:
            raise ProductNotFoundError(f'Price ID not found for product {product}')
        return price_id
    
    def get_price_for_price_per_foot(self, product, finish, width, height, with_damper=False, special_color_multiplier=None, price_id=None, height_unit='inches'):
        """Get price for a product with price_per_foot pricing
        
        Formula: (height_in_ft) × price_per_foot (matching width first)
        Height is converted to feet based on height_unit:
        - If unit is mm, cm, or m: convert back to original unit, then to meters, then to ft using M * 3.281
        - Otherwise (inches): use height / 12
        
        Args:
            product: Product model name
            finish: Finish type
            width: Width in inches (must match database width) - not used if price_id is provided
            height: Height in inches
            with_damper: Whether product has damper option
            special_color_multiplier: Multiplier for special color finish
            price_id: Optional price_id to use directly (for has_no_dimensions case)
            height_unit: Unit of the original height dimension ('mm', 'cm', 'm', 'inches', etc.)
            
        Returns:
            Tuple of (calculated price, finish_multiplier)
            
        Raises:
            ProductNotFoundError: If product data not found
            PriceNotFoundError: If price_per_foot not found
        """
        # Load product data
        table_id, base_modifier, anodized_multiplier, powder_coated_multiplier, no_finish_multiplier, wd_modifier = self._load_product_data(product)
        
        # Get price_per_foot
        # Note: 'width' parameter is actually the size value stored in SQL height column
        if price_id is not None:
            price_per_foot = self.db._get_price_per_unit(table_id, 0, price_id, 'price_per_foot')
        else:
            price_per_foot = self.db._get_price_per_unit(table_id, width, None, 'price_per_foot')
        
        if price_per_foot is None:
            raise PriceNotFoundError(f'Price per foot not found for product {product}')
        
        # Convert height to feet based on unit
        # height is currently in inches, so we need to convert it back to original unit first
        unit_lower = str(height_unit).lower().strip()
        
        if 'mm' in unit_lower or 'millimeter' in unit_lower:
            # Convert inches -> mm -> m -> ft
            # inches to mm: multiply by 25
            # mm to m: divide by 1000
            # m to ft: multiply by 3.281
            height_in_ft = (height * 25 / 1000) * 3.281
        elif 'cm' in unit_lower or 'centimeter' in unit_lower:
            # Convert inches -> cm -> m -> ft
            # inches to cm: multiply by 2.5
            # cm to m: divide by 100
            # m to ft: multiply by 3.281
            height_in_ft = (height * 2.5 / 100) * 3.281
        elif unit_lower == 'm' or 'meter' in unit_lower:
            # Convert inches -> m -> ft
            # inches to m: divide by 40
            # m to ft: multiply by 3.281
            height_in_ft = (height / 40) * 3.281
        else:
            # Default: inches -> ft (divide by 12)
            height_in_ft = height / 12
        
        # Calculate base price: height_in_ft * price_per_foot
        # Note: 'height' parameter is the dimension to multiply with price_per_foot
        # For price_per_foot, we use the calculated base_price as both TB and WD
        tb_price = height_in_ft * price_per_foot
        wd_price = tb_price  # Price per foot products use the same base for both
        
        # Use shared helper function to calculate final price
        final_price, finish_multiplier = self._calculate_final_price_from_base_prices(
            tb_price, wd_price, base_modifier, anodized_multiplier,
            powder_coated_multiplier, no_finish_multiplier, wd_modifier, finish, with_damper,
            special_color_multiplier
        )
        return final_price, finish_multiplier
    
    def get_price_for_price_per_sq_in(self, product, finish, db_size, actual_width, actual_height, with_damper=False, special_color_multiplier=None, price_id=None, width_unit='inches', height_unit='inches'):
        """Get price for a product with price_per_sq_in pricing
        
        Formula: (actual_width * actual_height) × price_per_sq_in (matching db_size first)
        Both actual_width and actual_height are in inches.
        
        Args:
            product: Product model name
            finish: Finish type
            db_size: Size value stored in SQL height column (used to look up price_per_sq_in) - not used if price_id is provided
            actual_width: Actual width in inches (for area calculation)
            actual_height: Actual height in inches (for area calculation)
            with_damper: Whether product has damper option
            special_color_multiplier: Multiplier for special color finish
            price_id: Optional price_id to use directly (for has_no_dimensions case)
            width_unit: Unit of the original width dimension ('mm', 'cm', 'm', 'inches', etc.) - not used currently
            height_unit: Unit of the original height dimension ('mm', 'cm', 'm', 'inches', etc.) - not used currently
            
        Returns:
            Tuple of (calculated price, finish_multiplier)
            
        Raises:
            ProductNotFoundError: If product data not found
            PriceNotFoundError: If price_per_sq_in not found
        """
        # Load product data
        table_id, base_modifier, anodized_multiplier, powder_coated_multiplier, no_finish_multiplier, wd_modifier = self._load_product_data(product)
        
        # Get price_per_sq_in
        # Note: 'db_size' parameter is the size value stored in SQL height column
        if price_id is not None:
            price_per_sq_in = self.db._get_price_per_unit(table_id, 0, price_id, 'price_per_sq_in')
        else:
            price_per_sq_in = self.db._get_price_per_unit(table_id, db_size, None, 'price_per_sq_in')
        
        if price_per_sq_in is None:
            raise PriceNotFoundError(f'Price per sq.in. not found for product {product}')
        
        # Calculate area in square inches: actual_width * actual_height
        # Both actual_width and actual_height are already in inches
        area_sq_in = actual_width * actual_height
        
        # Calculate base price: area_sq_in * price_per_sq_in
        # For price_per_sq_in, we use the calculated base_price as both TB and WD
        tb_price = area_sq_in * price_per_sq_in
        wd_price = tb_price  # Price per sq.in. products use the same base for both
        
        # Use shared helper function to calculate final price
        final_price, finish_multiplier = self._calculate_final_price_from_base_prices(
            tb_price, wd_price, base_modifier, anodized_multiplier,
            powder_coated_multiplier, no_finish_multiplier, wd_modifier, finish, with_damper,
            special_color_multiplier
        )
        return final_price, finish_multiplier
    
    def find_rounded_price_per_foot_width(self, product, width):
        """Find the exact match first, then the next available width that is >= the given width for price_per_foot products
        
        Returns:
            Rounded width
            
        Raises:
            SizeNotFoundError: If rounded width not found
        """
        rounded_width = self.db._find_rounded_price_per_unit_width(product, width, 'price_per_foot')
        if rounded_width is None:
            raise SizeNotFoundError(f'Height {width}" not available in price list for {product}. Please check available heights in the database.')
        return rounded_width
    
    def find_rounded_price_per_sq_in_width(self, product, width):
        """Find the exact match first, then the next available width that is >= the given width for price_per_sq_in products
        
        Returns:
            Rounded width
            
        Raises:
            SizeNotFoundError: If rounded width not found
        """
        rounded_width = self.db._find_rounded_price_per_unit_width(product, width, 'price_per_sq_in')
        if rounded_width is None:
            raise SizeNotFoundError(f'Height {width}" not available in price list for {product}. Please check available heights in the database.')
        return rounded_width
    
    def get_price_for_default_table(self, product, finish, size, with_damper=False, special_color_multiplier=None):
        """Get price for a specific product configuration using idx_price_lookup index
        
        Returns:
            Tuple of (calculated price, finish_multiplier)
        
        Raises:
            ValueError: If size format is invalid
            ProductNotFoundError: If product data not found
            PriceNotFoundError: If price not found
        """
        # Validate size input
        if not size:
            raise ValueError(f'Size cannot be None or empty')
        
        # Parse size to get width and height (supports decimal values)
        size_match = re.search(r'(\d+(?:\.\d+)?)"?\s*x\s*(\d+(?:\.\d+)?)"?', str(size).lower())
        if not size_match:
            raise ValueError(f'Invalid size format: {size}')
        
        width = float(size_match.group(1))
        height = float(size_match.group(2))
        
        # Load product data
        table_id, base_modifier, anodized_multiplier, powder_coated_multiplier, no_finish_multiplier, wd_modifier = self._load_product_data(product)
        
        # Special case for Model VD, VD-G and VD-M: Check for oversized dimensions first
        if product in ['VD', 'VD-G', 'VD-M']:
            # Check for VD oversized dimensions using inch values directly
            vd_result = self._calculate_vd_oversized_price(table_id, width, height, with_damper, product)
            if vd_result is not None:
                # vd_result contains (tb_price, wd_price) after dimension multipliers
                tb_price, wd_price = vd_result
                # Apply finish multipliers using the shared helper function
                final_price, finish_multiplier = self._calculate_final_price_from_base_prices(
                    tb_price, wd_price, base_modifier, anodized_multiplier,
                    powder_coated_multiplier, no_finish_multiplier, wd_modifier, finish, with_damper,
                    special_color_multiplier
                )
                return final_price, finish_multiplier
        
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
                raise PriceNotFoundError(f'Price not found for {product} with size {size}')
            
            # Get base prices (tb_price, wd_price)
            tb_price, wd_price = price_result
        
        # Use shared helper function to calculate final price
        final_price, finish_multiplier = self._calculate_final_price_from_base_prices(
            tb_price, wd_price, base_modifier, anodized_multiplier,
            powder_coated_multiplier, no_finish_multiplier, wd_modifier, finish, with_damper,
            special_color_multiplier
        )
        
        return final_price, finish_multiplier
    
    def find_rounded_default_table_size(self, product, finish, width, height):
        """Find the exact match first, then the next available size that is >= the given width and height
        
        Returns:
            Rounded size string or None if not found (None is acceptable here as it's used with 'or' operator)
        """
        return self.db.find_rounded_default_table_size(product, width, height)
    
    def get_price_for_other_table(self, product, finish, diameter, with_damper=False, special_color_multiplier=None):
        """Get price for an other table (diameter-based) product configuration
        
        Returns:
            Tuple of (calculated price, finish_multiplier)
        
        Raises:
            ValueError: If diameter format is invalid
            ProductNotFoundError: If product data not found
            PriceNotFoundError: If price not found
        """
        # Parse diameter value if it's a string (e.g., "8\" diameter" -> 8.0, "7.2\" diameter" -> 7.2)
        if isinstance(diameter, str):
            diameter_match = re.search(r'(\d+(?:\.\d+)?)', diameter)
            if not diameter_match:
                raise ValueError(f'Invalid diameter format: {diameter}')
            diameter = float(diameter_match.group(1))
        
        # Load product data
        table_id, base_modifier, anodized_multiplier, powder_coated_multiplier, no_finish_multiplier, wd_modifier = self._load_product_data(product)
        
        # Query prices using the table_id and diameter
        price_result = self.db.get_price_for_diameter(table_id, diameter)
        if not price_result:
            raise PriceNotFoundError(f'Price not found for {product} with diameter {diameter}"')
        
        # Get base prices (tb_price, wd_price)
        tb_price, wd_price = price_result
        
        # Use shared helper function to calculate final price
        final_price, finish_multiplier = self._calculate_final_price_from_base_prices(
            tb_price, wd_price, base_modifier, anodized_multiplier,
            powder_coated_multiplier, no_finish_multiplier, wd_modifier, finish, with_damper,
            special_color_multiplier
        )
        return final_price, finish_multiplier
    
    def find_rounded_other_table_size(self, product, diameter, price_id=None):
        """Find the exact match first, then the next available diameter that is >= the given diameter for other table products
        
        Args:
            product: Product model name
            diameter: Diameter in inches (not used if price_id is provided)
            price_id: Optional price_id to use directly (for has_no_dimensions case)
            
        Returns:
            Size string (e.g., "8\" diameter")
            
        Raises:
            SizeNotFoundError: If rounded size not found
        """
        # Convert diameter to float (supports decimal values)
        if isinstance(diameter, str):
            diameter_match = re.search(r'(\d+(?:\.\d+)?)', diameter)
            if not diameter_match:
                raise SizeNotFoundError(f'Invalid diameter format: {diameter}')
            diameter_float = float(diameter_match.group(1))
        else:
            diameter_float = float(diameter)
        rounded_size = self.db.find_rounded_other_table_size(product, diameter_float, price_id)
        if not rounded_size:
            raise SizeNotFoundError(f'Size not available for {product}')
        return rounded_size
    
    def has_damper_option(self, product):
        """Check if a product has a non-null WD multiplier in the header sheet"""
        return self.db.has_damper_option(product)
    
    def _calculate_vd_oversized_price(self, table_id, width, height, with_damper=False, product=None):
        """
        Calculate base TB and WD prices for VD, VD-G and VD-M products when dimensions exceed limits:
        - Height > 40 inch or Width > 80 inch
        - Use maximum dimensions (40 inch height, 80 inch width) to calculate round-up multipliers
        - Find price for the divided dimensions and multiply by round-up numbers
        - Returns base prices (TB and WD) after dimension multipliers, ready for finish multiplier application
        
        Args:
            table_id: Table ID for database lookup
            width: Width in inches
            height: Height in inches
            with_damper: Whether product has damper option
            product: Product model name (VD, VD-G or VD-M) - kept for compatibility but not used
            
        Returns:
            Tuple of (tb_price, wd_price) after dimension multipliers, or None if not oversized
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
        
        # Keep adjusted dimensions as floats for database lookup (supports decimal values)
        adjusted_height_inches = adjusted_height
        adjusted_width_inches = adjusted_width
        
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
                lookup_height, lookup_width, tb_price, wd_price = closest
            else:
                return None
        else:
            # Get both TB and WD prices from exact match
            tb_price, wd_price = price_result
        
        # Calculate total multiplier
        total_multiplier = height_multiplier * width_multiplier
        
        # Apply dimension multipliers to both TB and WD prices
        # These will be passed to _calculate_final_price_from_base_prices for finish multiplier application
        tb_price_with_multiplier = tb_price * total_multiplier
        wd_price_with_multiplier = wd_price * total_multiplier
        
        return tb_price_with_multiplier, wd_price_with_multiplier
    
    def __del__(self):
        """Clean up database connection when object is destroyed"""
        if hasattr(self, 'db'):
            self.db.close()
