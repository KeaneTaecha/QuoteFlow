"""
Quote Utilities
Shared functions for building quote items with pricing calculations.
"""

import re
from typing import Dict, Optional, Tuple
from utils.price_calculator import PriceCalculator, PriceNotFoundError, ProductNotFoundError, SizeNotFoundError
from utils.filter_utils import get_filter_price
from utils.product_utils import convert_dimension_to_inches


def calculate_ins_price(width_inches: float, height_inches: float) -> float:
    """Calculate INS price based on square inches. Minimum price is 50."""
    return max(width_inches * height_inches * 0.15, 50.0)


def _build_original_size(width, height, width_unit, height_unit, 
                         slot_number, rounded_size, has_price_per_foot, is_other_table, has_no_dimensions):
    """Build original size string from dimensions, preserving original units."""
    def format_dimension(value, unit):
        """Format a dimension value with its unit symbol."""
        unit_lower = str(unit).lower().strip()
        if unit_lower in ['millimeters', 'mm']:
            return f"{value}mm"
        elif unit_lower in ['centimeters', 'cm']:
            return f"{value}cm"
        elif unit_lower in ['meters', 'm']:
            return f"{value}m"
        elif unit_lower in ['feet', 'ft', "'"]:
            return f"{value}ft"
        elif unit_lower in ['inches', 'inch', 'in', '"']:
            return f'{value}"'
        else:
            # Default to inches format
            return f'{value}"'
    
    if has_no_dimensions:
        if has_price_per_foot and height is not None:
            height_str = format_dimension(height, height_unit)
            return f"{slot_number}Slot x {height_str}" if slot_number else f"Slot x {height_str}"
        elif rounded_size:
            # For non-price_per_foot has_no_dimensions products, treat as diameter-based
            # rounded_size is already in inches format from database
            diameter_match = re.search(r'(\d+(?:\.\d+)?)', rounded_size)
            if diameter_match:
                diameter = diameter_match.group(1)
                return f"{slot_number}Slot x {diameter}\"" if slot_number else f"Slot x {diameter}\""
        return None
    
    if is_other_table and height is not None:
        # For other_table products, height represents the diameter
        return format_dimension(height, height_unit)
    elif width and height:
        width_str = format_dimension(width, width_unit)
        height_str = format_dimension(height, height_unit)
        return f'{width_str} x {height_str}'
    return None


def _calculate_filter_and_ins(price_calculator, filter_type, has_ins, width_inches=None, height_inches=None):
    """Calculate filter and INS prices for default table products only.
    
    Args:
        price_calculator: PriceCalculator instance
        filter_type: Optional filter type (e.g., "Nylon")
        has_ins: Whether product has INS option
        width_inches: Width in inches (required for filter calculation)
        height_inches: Height in inches (required for filter calculation)
    
    Raises:
        ValueError: If filter not found in database or dimensions are missing
    """
    filter_price = 0.0
    if filter_type:
        if width_inches is not None and height_inches is not None:
            filter_price = get_filter_price(price_calculator, filter_type, max(width_inches, height_inches), width_inches, height_inches)
            if filter_price is None:
                raise ValueError(f'Filter "{filter_type}" not found in database')
        else:
            raise ValueError('Width and height are required for filter calculation')
    
    ins_price = calculate_ins_price(width_inches, height_inches) if has_ins and width_inches and height_inches else 0.0
    return filter_price, ins_price


def build_quote_item(
    price_calculator: PriceCalculator,
    product: str,
    finish: Optional[str],
    quantity: int,
    has_wd: bool,
    has_price_per_foot: bool,
    is_other_table: bool,
    width: Optional[float] = None,
    height: Optional[float] = None,
    width_unit: str = 'inches',
    height_unit: str = 'inches',
    filter_type: Optional[str] = None,
    discount: float = 0.0,  # Discount as percentage (0-100)
    special_color_multiplier: Optional[float] = None,
    product_code: Optional[str] = None,
    detail: str = '',
    has_ins: bool = False,
    has_no_dimensions: bool = False,
    slot_number: Optional[str] = None
) -> Tuple[Optional[Dict], Optional[str]]:
    """
    Build a quote item with pricing calculations.
    
    Args:
        price_calculator: PriceCalculator instance
        product: Base product name (without WD suffix)
        finish: Finish name
        quantity: Quantity
        has_wd: Whether product has WD (damper) option
        has_price_per_foot: Whether product uses price_per_foot pricing
        is_other_table: Whether product uses other_table (diameter-based) pricing
        width: Width dimension (for price_per_foot or default products)
        height: Height dimension (for price_per_foot, default products, or other_table products as diameter)
        width_unit: Unit for width ('inches' or 'millimeters')
        height_unit: Unit for height ('inches' or 'millimeters')
        filter_type: Optional filter type (e.g., "Nylon")
        discount: Discount percentage (0-100)
        special_color_multiplier: Multiplier for special color pricing (default None, uses 1.45 if not provided)
        product_code: Optional pre-built product code (if None, will be built)
        detail: Detail text for the item
        has_ins: Whether product has INS option (adds price based on square inches)
        has_no_dimensions: Whether product has no height and width in database
        slot_number: Optional slot number extracted from model name (for no-dimension products)
        
    Returns:
        Tuple of (quote_item_dict, error_message)
        If successful: (quote_item_dict, None)
        If error: (None, error_message_string)
    """
    warning_message = None
    table_price = price_after_finish = ins_price = filter_price = 0.0
    rounded_size = None
    finish_multiplier = None
    
    # If height is not provided but width is, use width as fallback with warning
    if height is None and width is not None:
        height = width
        height_unit = width_unit
        warning_message = f'Width was used as height (diameter) for other_table product {product}. Please use height field for diameter.'

    # Handle no_dimensions products
    if has_no_dimensions:
        try:
            price_id = price_calculator.get_price_id_for_no_dimensions(product)
        except ProductNotFoundError as e:
            return None, str(e)
        
        if not has_price_per_foot:
            return None, f'Product {product} has no dimensions but does not use price_per_foot pricing'
        
        if has_price_per_foot:
            if height is None:
                return None, f'Height is required for price_per_foot product {product}'
            try:
                height_inches = convert_dimension_to_inches(height, height_unit)
            except ValueError as e:
                return None, str(e)
            try:
                table_price, _ = price_calculator.get_price_for_price_per_foot(
                    product, None, 0, height_inches, has_wd, 1.0, price_id=price_id, height_unit=height_unit
                )
            except (ProductNotFoundError, PriceNotFoundError) as e:
                return None, str(e)
            
            try:
                price_after_finish, finish_multiplier = price_calculator.get_price_for_price_per_foot(
                    product, finish, 0, height_inches, has_wd, special_color_multiplier, price_id=price_id, height_unit=height_unit
                )
            except (ProductNotFoundError, PriceNotFoundError) as e:
                return None, str(e)
            
            # INS and filter calculations only apply to default table products
            # Skip for has_no_dimensions products
            filter_price, ins_price = 0.0, 0.0
            
        
    # Handle price_per_foot products
    elif has_price_per_foot:
        if width is None or height is None:
            return None, f'Width and height required for price_per_foot product {product}'
        
        try:
            width_inches = convert_dimension_to_inches(width, width_unit)
            height_inches = convert_dimension_to_inches(height, height_unit)
        except ValueError as e:
            return None, str(e)
        
        if height_inches > width_inches:
            width_inches, height_inches = height_inches, width_inches
            warning_message = f'Width and height appear to be swapped. Using {width_inches}" x {height_inches}" instead.'
        
        try:
            rounded_height = price_calculator.find_rounded_price_per_foot_width(product, height_inches)
        except SizeNotFoundError as e:
            return None, str(e)
        
        try:
            # width_inches is passed as 'height' parameter (dimension to multiply), so use width_unit
            table_price, _ = price_calculator.get_price_for_price_per_foot(product, None, rounded_height, width_inches, has_wd, 1.0, height_unit=width_unit)
        except (ProductNotFoundError, PriceNotFoundError) as e:
            return None, str(e)
        
        try:
            # width_inches is passed as 'height' parameter (dimension to multiply), so use width_unit
            price_after_finish, finish_multiplier = price_calculator.get_price_for_price_per_foot(
                product, finish, rounded_height, width_inches, has_wd, special_color_multiplier, height_unit=width_unit
            )
        except (ProductNotFoundError, PriceNotFoundError) as e:
            return None, str(e)
        
        # INS and filter calculations only apply to default table products
        # Skip for has_price_per_foot products
        filter_price, ins_price = 0.0, 0.0
        
        rounded_size = f'{width_inches}" x {rounded_height}"'
        
    # Handle other_table products
    elif is_other_table:
        # Height is already validated and set at the beginning of the function
        try:
            height_inches = convert_dimension_to_inches(height, height_unit)
        except ValueError as e:
            return None, str(e)
        
        try:
            rounded_size = price_calculator.find_rounded_other_table_size(product, height_inches)
        except SizeNotFoundError:
            rounded_size = f'{height_inches}" diameter'
        
        try:
            table_price, _ = price_calculator.get_price_for_other_table(product, None, rounded_size, has_wd, 1.0)
        except (ProductNotFoundError, PriceNotFoundError, ValueError) as e:
            return None, str(e)
        
        try:
            price_after_finish, finish_multiplier = price_calculator.get_price_for_other_table(product, finish, rounded_size, has_wd, special_color_multiplier)
        except (ProductNotFoundError, PriceNotFoundError, ValueError) as e:
            return None, str(e)
        
        # INS and filter calculations only apply to default table products
        # Skip for is_other_table products
        filter_price, ins_price = 0.0, 0.0
        
    # Handle default width/height products
    else:
        # If only one dimension is provided, apply it to both with warning
        if width is not None and height is None:
            height = width
            height_unit = width_unit
            warning_message = f'Only width ({width} {width_unit}) was provided for default table. Applied to both height and width.'
        elif height is not None and width is None:
            width = height
            width_unit = height_unit
            warning_message = f'Only height ({height} {height_unit}) was provided for default table. Applied to both height and width.'
        
        if width is None or height is None:
            return None, f'Width and height required for {product}'
        
        try:
            width_inches = convert_dimension_to_inches(width, width_unit)
            height_inches = convert_dimension_to_inches(height, height_unit)
        except ValueError as e:
            return None, str(e)
        
        if height_inches > width_inches:
            width_inches, height_inches = height_inches, width_inches
            width, height = height, width
            width_unit, height_unit = height_unit, width_unit
            warning_message = f'Width and height appear to be swapped. Using {width_inches}" x {height_inches}" instead.'
        
        rounded_size = price_calculator.find_rounded_default_table_size(product, finish, width_inches, height_inches) or f'{width_inches}" x {height_inches}"'
        
        try:
            table_price, _ = price_calculator.get_price_for_default_table(product, None, rounded_size, has_wd, 1.0)
        except (ProductNotFoundError, PriceNotFoundError, ValueError) as e:
            return None, str(e)
        
        try:
            price_after_finish, finish_multiplier = price_calculator.get_price_for_default_table(product, finish, rounded_size, has_wd, special_color_multiplier)
        except (ProductNotFoundError, PriceNotFoundError, ValueError) as e:
            return None, str(e)
        
        try:
            filter_price, ins_price = _calculate_filter_and_ins(price_calculator, filter_type, has_ins, width_inches, height_inches)
        except ValueError as e:
            return None, str(e)
        
        # Add handgear addition to ins_price for VD, VD-G, and VD-M products
        # This allows handgear to be displayed separately in the INS column in Excel
        # Note: Handgear is no longer included in price_after_finish (removed from get_price_for_default_table)
        if product in ['VD-G', 'VD-M', 'RVD-G', 'RVD-M']:
            hand_gear_addition = price_calculator.get_hand_gear_price(product, width_inches, height_inches)
            ins_price = ins_price + hand_gear_addition
    
    # Build original size
    original_size = _build_original_size(width, height, width_unit, height_unit, 
                                          slot_number, rounded_size, has_price_per_foot, is_other_table, has_no_dimensions)
    
    # Calculate final prices
    # Note: finish_multiplier is already obtained from the price calculation above
    unit_price = price_after_finish + filter_price + ins_price
    discount_decimal = discount / 100.0
    discounted_unit_price = unit_price * (1 - discount_decimal)
    
    # Build product code if not provided
    if product_code is None:
        product_code = product
        if has_wd:
            product_code = f"{product_code}(WD)"
        if has_ins:
            product_code = f"{product_code}(INS)"
        if filter_type:
            product_code = f"{product_code}+F.{filter_type}"
    
    # Build quote item
    quote_item = {
        'product_code': product_code,
        'size': original_size or '',
        'finish': finish,
        'quantity': quantity,
        'unit_price': unit_price,
        'discount': discount_decimal,
        'discounted_unit_price': discounted_unit_price,
        'total': int((discounted_unit_price * quantity) + 0.5),
        'rounded_size': rounded_size,
        'detail': detail,
        'table_price': table_price,
        'price_after_finish': price_after_finish,
        'ins_price': ins_price,
        'filter_price': filter_price,
        'finish_multiplier': finish_multiplier
    }
    
    if warning_message:
        quote_item['warning_message'] = warning_message
    
    return quote_item, None

