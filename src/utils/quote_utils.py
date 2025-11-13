"""
Quote Utilities
Shared functions for building quote items with pricing calculations.
"""

import re
from typing import Dict, Optional, Tuple
from utils.price_calculator import PriceCalculator
from utils.filter_utils import get_filter_price
from utils.product_utils import convert_dimension_to_inches


def calculate_ins_price(width_inches: float, height_inches: float) -> float:
    """
    Calculate INS price based on square inches.
    
    INS price is calculated as: (width × height) × 0.15
    Minimum price is 50.
    
    Args:
        width_inches: Width in inches
        height_inches: Height in inches
        
    Returns:
        INS price (minimum 50)
    """
    square_inches = width_inches * height_inches
    ins_price = square_inches * 0.15
    return max(ins_price, 50.0)  # Minimum price is 50


def build_quote_item(
    price_loader: PriceCalculator,
    product: str,
    finish: Optional[str],
    quantity: int,
    has_wd: bool,
    has_price_per_foot: bool,
    is_other_table: bool,
    width: Optional[float] = None,
    height: Optional[float] = None,
    size: Optional[float] = None,
    width_unit: str = 'inches',
    height_unit: str = 'inches',
    size_unit: str = 'inches',
    filter_type: Optional[str] = None,
    discount: float = 0.0,  # Discount as percentage (0-100)
    special_color_multiplier: float = 1.0,
    product_code: Optional[str] = None,
    original_size: Optional[str] = None,
    detail: str = '',
    has_ins: bool = False,
    has_no_dimensions: bool = False,
    slot_number: Optional[str] = None
) -> Tuple[Optional[Dict], Optional[str]]:
    """
    Build a quote item with pricing calculations.
    
    Args:
        price_loader: PriceCalculator instance
        product: Base product name (without WD suffix)
        finish: Finish name
        quantity: Quantity
        has_wd: Whether product has WD (damper) option
        has_price_per_foot: Whether product uses price_per_foot pricing
        is_other_table: Whether product uses other_table (diameter-based) pricing
        width: Width dimension (for price_per_foot or default products)
        height: Height dimension (for price_per_foot or default products)
        size: Size dimension (for other_table products)
        width_unit: Unit for width ('inches' or 'millimeters')
        height_unit: Unit for height ('inches' or 'millimeters')
        size_unit: Unit for size ('inches' or 'millimeters')
        filter_type: Optional filter type (e.g., "Nylon")
        discount: Discount percentage (0-100)
        special_color_multiplier: Multiplier for special color pricing (default 1.0)
        product_code: Optional pre-built product code (if None, will be built)
        original_size: Optional pre-formatted original size string (if None, will be built)
        detail: Detail text for the item
        has_ins: Whether product has INS option (adds price based on square inches)
        has_no_dimensions: Whether product has no height and width in database
        slot_number: Optional slot number extracted from model name (for no-dimension products)
        
    Returns:
        Tuple of (quote_item_dict, error_message)
        If successful: (quote_item_dict, None)
        If error: (None, error_message_string)
    """
    warning_message = None  # Initialize warning message
    
    if has_no_dimensions:
        # Handle products with no height/width - extract price_id first
        price_id = price_loader.get_price_id_for_no_dimensions(product)
        if price_id is None:
            return None, f'Price ID not found for product {product}'
        
        # Use price_id with appropriate function based on product type
        if has_price_per_foot:
            # For price_per_foot products with no dimensions, use get_price_for_price_per_foot with price_id
            # Note: height is still required for price_per_foot calculation
            if height is None:
                return None, f'Height is required for price_per_foot product {product}'
            
            height_inches = convert_dimension_to_inches(height, height_unit)
            unit_price = price_loader.get_price_for_price_per_foot(
                product, finish, 0, height_inches, has_wd, special_color_multiplier, price_id=price_id
            )
            if unit_price == 0:
                return None, f'Price not available for {product}'
            
            rounded_size = None
        elif is_other_table:
            # For other table products with no dimensions, use find_rounded_other_table_size with price_id
            rounded_size = price_loader.find_rounded_other_table_size(product, finish, 0, price_id=price_id)
            if not rounded_size:
                return None, f'Size not available for {product}'
            
            # Get price using the rounded size
            unit_price = price_loader.get_price_for_other_table(product, finish, rounded_size, has_wd, special_color_multiplier)
            if unit_price == 0:
                return None, f'Price not available for {product} {rounded_size}'
        else:
            # Fallback: if neither price_per_foot nor other_table, return error
            return None, f'Product {product} has no dimensions but is not price_per_foot or other_table'
        
        # Set size to "[number]Slot x Height" format for no-dimension products
        if has_price_per_foot and height is not None:
            # For price_per_foot products, we have height
            # Preserve the original unit (mm or inches)
            if height_unit.lower() in ['millimeters', 'mm']:
                if slot_number:
                    original_size = f"{slot_number}Slot x {height}mm"
                else:
                    original_size = f"Slot x {height}mm"
            else:
                # height_unit is inches
                height_inches = convert_dimension_to_inches(height, height_unit)
                if slot_number:
                    original_size = f"{slot_number}Slot x {height_inches}\""
                else:
                    original_size = f"Slot x {height_inches}\""
        elif is_other_table and rounded_size:
            # For other_table products, extract diameter from rounded_size (e.g., "8\" diameter" -> "8")
            # Use diameter as "height" in the format
            diameter_match = re.search(r'(\d+(?:\.\d+)?)', rounded_size)
            if diameter_match:
                diameter_value = diameter_match.group(1)
                if slot_number:
                    original_size = f"{slot_number}Slot x {diameter_value}\""
                else:
                    original_size = f"Slot x {diameter_value}\""
            else:
                # Cannot extract diameter from rounded_size
                return None, f'Price not available for {product}: cannot extract diameter from size'
        else:
            # No height available for no-dimension product
            return None, f'Price not available for {product}: height is required'
        
        # No filter or INS price for no-dimension products (they don't have dimensions)
        
    elif has_price_per_foot:
        # Handle price_per_foot products - require width and height
        if width is None or height is None:
            return None, f'Width and height required for price_per_foot product {product}'
        
        # Convert to inches if needed
        width_inches = convert_dimension_to_inches(width, width_unit)
        height_inches = convert_dimension_to_inches(height, height_unit)
        
        # Check if height is greater than width - if so, swap them and show warning
        if height_inches > width_inches:
            # Swap the values and add warning
            width_inches, height_inches = height_inches, width_inches
            warning_message = f'Width and height appear to be swapped. Using {width_inches}" x {height_inches}" instead.'
            # Store warning message separately (will be displayed like error messages)
            # Don't add to detail field - it will be shown in product column
        
        # Find the rounded height that matches database (for other tables, size is stored in height column)
        height_int = int(round(height_inches))
        rounded_height = price_loader.find_rounded_price_per_foot_width(product, height_int)
        
        if not rounded_height:
            return None, f'Height {height_inches}" ({height_int}") not available in price list for {product}. Please check available heights in the database.'
        
        # Get price using price_per_foot formula
        unit_price = price_loader.get_price_for_price_per_foot(
            product, finish, rounded_height, width_inches, has_wd, special_color_multiplier
        )
        if unit_price == 0:
            return None, f'Price not available for {product}'
        
        # Add filter price if filter is specified
        if filter_type:
            filter_price = get_filter_price(
                price_loader, filter_type, max(width_inches, height_inches), width_inches, height_inches
            )
            if filter_price is None:
                return None, f'Filter "{filter_type}" not found in database'
            unit_price += filter_price
        
        # Add INS price if INS is specified
        if has_ins:
            ins_price = calculate_ins_price(width_inches, height_inches)
            unit_price += ins_price
        
        # Build original size if not provided
        if original_size is None:
            if width_unit == 'millimeters' and height_unit == 'millimeters':
                original_size = f"{width}mm x {height}mm"
            elif width_unit == 'inches' and height_unit == 'inches':
                original_size = f'{width}" x {height}"'
            else:
                # Mixed units
                width_str = f"{width}mm" if width_unit == 'millimeters' else f'{width}"'
                height_str = f"{height}mm" if height_unit == 'millimeters' else f'{height}"'
                original_size = f'{width_str} x {height_str}'
        
        # Store rounded size for display
        rounded_size = f'{width_inches}" x {rounded_height}"'
        
    elif is_other_table:
        # Handle diameter-based products
        if size is None:
            return None, f'Size required for other_table product {product}'
        
        # Convert to inches if needed
        size_inches = convert_dimension_to_inches(size, size_unit)
        
        # Find rounded size and get price
        rounded_size = price_loader.find_rounded_other_table_size(product, finish, size_inches)
        if not rounded_size:
            # Try direct lookup for exceeded dimensions
            unit_price = price_loader.get_price_for_other_table(
                product, finish, f'{size_inches}" diameter', has_wd, special_color_multiplier
            )
            if unit_price == 0:
                return None, f'Size not available for {product}'
            rounded_size = f'{size_inches}" diameter'
        else:
            unit_price = price_loader.get_price_for_other_table(
                product, finish, rounded_size, has_wd, special_color_multiplier
            )
            if unit_price == 0:
                return None, f'Price not available for {product} {rounded_size}'
        
        # Add filter price if filter is specified
        if filter_type:
            filter_price = get_filter_price(price_loader, filter_type, size_inches)
            if filter_price is None:
                return None, f'Filter "{filter_type}" not found in database'
            unit_price += filter_price
        
        # Build original size if not provided
        if original_size is None:
            if size_unit == 'millimeters':
                original_size = f"{size}mm"
            else:
                original_size = f'{size}"'
        
    else:
        # Handle width/height-based products
        if width is None or height is None:
            return None, f'Width and height required for {product}'
        
        # Convert to inches if needed
        width_inches = convert_dimension_to_inches(width, width_unit)
        height_inches = convert_dimension_to_inches(height, height_unit)
        
        # Check if height is greater than width - if so, swap them and show warning
        if height_inches > width_inches:
            # Swap the values and add warning
            width_inches, height_inches = height_inches, width_inches
            # Also swap the original width/height and units for consistency
            width, height = height, width
            width_unit, height_unit = height_unit, width_unit
            warning_message = f'Width and height appear to be swapped. Using {width_inches}" x {height_inches}" instead.'
        
        # Find rounded size and get price
        rounded_size = price_loader.find_rounded_default_table_size(product, finish, width_inches, height_inches)
        if not rounded_size:
            # Try direct lookup for exceeded dimensions
            unit_price = price_loader.get_price_for_default_table(
                product, finish, f'{width_inches}" x {height_inches}"', has_wd, special_color_multiplier
            )
            if unit_price == 0:
                return None, f'Size not available for {product}'
            rounded_size = f'{width_inches}" x {height_inches}"'
        else:
            unit_price = price_loader.get_price_for_default_table(
                product, finish, rounded_size, has_wd, special_color_multiplier
            )
            if unit_price == 0:
                return None, f'Price not available for {product} {rounded_size}'
        
        # Add filter price if filter is specified
        if filter_type:
            filter_price = get_filter_price(
                price_loader, filter_type, max(width_inches, height_inches), width_inches, height_inches
            )
            if filter_price is None:
                return None, f'Filter "{filter_type}" not found in database'
            unit_price += filter_price
        
        # Add INS price if INS is specified
        if has_ins:
            ins_price = calculate_ins_price(width_inches, height_inches)
            unit_price += ins_price
        
        # Build original size if not provided
        if original_size is None:
            if width_unit == 'millimeters' and height_unit == 'millimeters':
                original_size = f"{width}mm x {height}mm"
            elif width_unit == 'inches' and height_unit == 'inches':
                original_size = f'{width}" x {height}"'
            else:
                # Mixed units
                width_str = f"{width}mm" if width_unit == 'millimeters' else f'{width}"'
                height_str = f"{height}mm" if height_unit == 'millimeters' else f'{height}"'
                original_size = f'{width_str} x {height_str}'
        
        # rounded_size is already set above
    
    # Build product code if not provided
    if product_code is None:
        product_code = product
        if has_wd:
            product_code = f"{product_code}(WD)"
        if has_ins:
            product_code = f"{product_code}(INS)"
        if filter_type:
            product_code = f"{product_code}+F.{filter_type}"
    
    # Apply discount
    discount_decimal = discount / 100.0  # Convert percentage to decimal
    discount_amount = unit_price * discount_decimal
    discounted_unit_price = unit_price - discount_amount
    total_price = discounted_unit_price * quantity
    
    # Build quote item
    quote_item = {
        'product_code': product_code,
        'size': original_size,
        'finish': finish,
        'quantity': quantity,
        'unit_price': unit_price,
        'discount': discount_decimal,  # Store as decimal (0.1 for 10%)
        'discounted_unit_price': discounted_unit_price,
        'total': total_price,
        'rounded_size': rounded_size,
        'detail': detail
    }
    
    # Add warning_message if warning exists (similar to error_message for invalid items)
    if warning_message:
        quote_item['warning_message'] = warning_message
    
    return quote_item, None

