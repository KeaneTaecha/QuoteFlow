"""
Product Utilities
Shared functions for product name extraction, unit conversion, and product validation.
"""

import re
from typing import Tuple, Optional, List
from utils.price_calculator import PriceCalculator


def extract_slot_number_from_model(model: str) -> Optional[str]:
    """
    Extract slot number from the beginning of model name.
    
    The slot number is the numeric prefix at the start of the model name.
    For example: "2ABC" -> "2", "10XYZ" -> "10", "ABC" -> None
    
    Args:
        model: Product model name
        
    Returns:
        Slot number as string (e.g., "2", "10") or None if no number found at start
    """
    if not model:
        return None
    
    # Match one or more digits at the start of the string
    match = re.match(r'^(\d+)', model.strip())
    if match:
        return match.group(1)
    
    return None


def extract_product_flags_and_filter(product_string: str) -> Tuple[str, bool, bool, Optional[str]]:
    """
    Extract base product name, WD flag, INS flag, and filter type from product string.
    Handles products with format: "ProductName(WD)(INS)+F.FilterType" or variations.
    
    Args:
        product_string: Product name that may include "(WD)" and/or "(INS)" suffix and/or "+F.xxx" filter suffix
        
    Returns:
        Tuple of (base_product_name, has_wd_flag, has_ins_flag, filter_type_or_none)
        
    Examples:
        "ProductName" -> ("ProductName", False, False, None)
        "ProductName(WD)" -> ("ProductName", True, False, None)
        "ProductName(INS)" -> ("ProductName", False, True, None)
        "ProductName(WD)(INS)" -> ("ProductName", True, True, None)
        "ProductName+F.Nylon" -> ("ProductName", False, False, "Nylon")
        "ProductName(WD)+F.Nylon" -> ("ProductName", True, False, "Nylon")
        "ProductName(INS)+F.Nylon" -> ("ProductName", False, True, "Nylon")
    """
    product = product_string.strip()
    has_wd = False
    has_ins = False
    filter_type = None

    # First extract filter if present (filter comes after everything)
    if '+F.' in product:
        parts = product.split('+F.', 1)
        if len(parts) == 2:
            product = parts[0].strip()
            filter_part = parts[1].strip()
            # Strip filter unit until it reaches (
            if '(' in filter_part:
                # Keep everything from ( onwards and append to product
                paren_index = filter_part.index('(')
                remaining_part = filter_part[paren_index:]
                product = product + remaining_part
                filter_type = filter_part[:paren_index].strip()
            else:
                filter_type = filter_part

    # Handle combined flags like "(WD,INS)", "(INS,WD)", or variants missing parentheses/spaces
    combined_flag_patterns = [
        r'\(?\s*WD\s*,\s*INS\s*\)?$',
        r'\(?\s*INS\s*,\s*WD\s*\)?$'
    ]
    combined_match_found = False
    for pattern in combined_flag_patterns:
        if re.search(pattern, product, flags=re.IGNORECASE):
            product = re.sub(pattern, '', product, flags=re.IGNORECASE).strip()
            combined_match_found = True
            break
    if combined_match_found:
        has_wd = True
        has_ins = True
    else:
        if "(INS)" in product:
            product = product.replace("(INS)", "").strip()
            has_ins = True
        
        if "(WD)" in product:
            product = product.replace("(WD)", "").strip()
            has_wd = True

    return product, has_wd, has_ins, filter_type


def convert_mm_to_inches(value_mm: float) -> float:
    """
    Convert millimeters to inches.
    
    Note: Uses division by 25.4 (standard conversion factor).
    
    Args:
        value_mm: Value in millimeters
        
    Returns:
        Value in inches
    """
    return value_mm / 25


def convert_cm_to_inches(value_cm: float) -> float:
    """
    Convert centimeters to inches.
    
    Args:
        value_cm: Value in centimeters
        
    Returns:
        Value in inches
    """
    return value_cm / 2.5


def convert_m_to_inches(value_m: float) -> float:
    """
    Convert meters to inches.
    
    Args:
        value_m: Value in meters
        
    Returns:
        Value in inches
    """
    return value_m * 40


def convert_ft_to_inches(value_ft: float) -> float:
    """
    Convert feet to inches.
    
    Args:
        value_ft: Value in feet
        
    Returns:
        Value in inches
    """
    return value_ft * 12


def convert_dimension_to_inches(value: float, unit: str) -> float:
    """
    Convert a dimension value to inches based on the unit.
    
    Args:
        value: The dimension value
        unit: Unit string ('Millimeters', 'millimeters', 'mm', 'Inches', 'inches', '"', 'cm', 'm', 'ft', etc.)
        
    Returns:
        Value in inches
        
    Raises:
        ValueError: If the unit is not recognized
    """
    unit_lower = str(unit).lower().strip()
    
    # Check for millimeters
    if 'mm' in unit_lower or 'millimeter' in unit_lower:
        return convert_mm_to_inches(value)
    
    # Check for centimeters
    if 'cm' in unit_lower or 'centimeter' in unit_lower:
        return convert_cm_to_inches(value)
    
    # Check for meters
    if unit_lower == 'm' or 'meter' in unit_lower:
        return convert_m_to_inches(value)
    
    # Check for feet
    if unit_lower == 'ft' or unit_lower == "'" or 'foot' in unit_lower or 'feet' in unit_lower:
        return convert_ft_to_inches(value)
    
    # Check for inches (including quote symbol)
    if 'inch' in unit_lower or unit_lower == '"' or unit_lower == 'in':
        return value
    
    # If unit is not recognized, raise an error
    raise ValueError(f'Incompatible unit "{unit}". Supported units are: inches (or "), millimeters (or mm), centimeters (or cm), meters (or m), feet (or ft)')


def find_matching_product(product: str, available_models: List[str], 
                          has_wd: Optional[bool] = None) -> Optional[Tuple[str, bool]]:
    """
    Find a matching product in the available models list using exact match only.
    
    Args:
        product: Product name to find
        available_models: List of available product models
        has_wd: Optional pre-extracted WD flag (avoids redundant extraction)
        
    Returns:
        Tuple of (matched_product_name, has_wd_flag) or None if not found
    """
    if not available_models:
        return None
    
    # Extract WD flag to return correct has_wd value
    if has_wd is None:
        _, has_wd, _, _ = extract_product_flags_and_filter(product)
    
    # Try exact match only
    if product in available_models:
        return product, has_wd
    
    return None


def validate_filter_exists(filter_type: str, price_calculator: PriceCalculator) -> bool:
    """
    Validate if a filter type exists in the database.
    
    Args:
        filter_type: Filter type to validate (e.g., "Nylon")
        price_calculator: PriceCalculator instance for database access
        
    Returns:
        True if filter exists, False otherwise
    """
    if not filter_type or not price_calculator:
        return False
    
    all_models = price_calculator.get_available_models()
    filter_type_lower = filter_type.lower()
    
    # Search for filter product that matches the filter type
    for model in all_models:
        if filter_type_lower in model.lower():
            return True
    
    return False


def get_product_type_flags(price_calculator: PriceCalculator, product: str) -> Tuple[bool, bool, bool, bool]:
    """
    Get product type flags (has_no_dimensions, has_price_per_foot, has_price_per_sq_in, is_other_table).
    Consolidates the logic for determining product type characteristics.
    
    Args:
        price_calculator: PriceCalculator instance
        product: Product model name
        
    Returns:
        Tuple of (has_no_dimensions, has_price_per_foot, has_price_per_sq_in, is_other_table)
    """
    has_no_dimensions = price_calculator.has_no_dimensions(product)
    has_price_per_foot = price_calculator.has_price_per_foot(product)
    has_price_per_sq_in = price_calculator.has_price_per_sq_in(product)
    
    # Determine is_other_table based on product characteristics
    # If has_no_dimensions is true, is_other_table must be false (no height = no diameter)
    if has_no_dimensions:
        is_other_table = False
    else:
        # Only check if it's NOT a price_per_foot or price_per_sq_in product
        is_other_table = price_calculator.is_other_table(product) if not (has_price_per_foot or has_price_per_sq_in) else False
    
    return has_no_dimensions, has_price_per_foot, has_price_per_sq_in, is_other_table


def parse_dimension_with_unit(value_str: str) -> Tuple[Optional[float], Optional[str]]:
    """
    Parse a dimension string and detect its unit.
    
    Detects units from the string itself. Returns full unit names:
    - 'millimeters' for mm
    - 'centimeters' for cm
    - 'meters' for m
    - 'feet' for ft
    - 'inches' for " or in
    
    Args:
        value_str: Dimension string (e.g., "1.0m", "550mm", "3\"")
        
    Returns:
        Tuple of (numeric_value, unit_name) or (None, None) if parsing fails.
        If no unit is detected, returns (numeric_value, None).
        
    Examples:
        - "1.0m" → (1.0, 'meters')
        - "550mm" → (550.0, 'millimeters')
        - "3\"" → (3.0, 'inches')
        - "2ft" → (2.0, 'feet')
        - "50cm" → (50.0, 'centimeters')
        - "550" → (550.0, None)
    """
    if not value_str:
        return None, None
    
    value_str = str(value_str).strip()
    value_lower = value_str.lower()
    
    # Extract numeric value first
    match = re.search(r'(\d+(?:\.\d+)?)', value_str)
    if not match:
        return None, None
    num_value = float(match.group(1))
    
    # Check for units in order of specificity (longer units first to avoid partial matches)
    # Check for quote (") - inches
    if '"' in value_str:
        return num_value, 'inches'
    
    # Check for single quote (') - feet
    if "'" in value_str:
        return num_value, 'feet'
    
    # Check for cm (but not part of mm) - check before mm
    if 'cm' in value_lower and 'mm' not in value_lower:
        return num_value, 'centimeters'
    
    # Check for mm
    if 'mm' in value_lower:
        return num_value, 'millimeters'
    
    # Check for meters - must check that 'm' is not part of 'mm' or 'cm'
    # Check if 'm' appears (as standalone or in 'meter'/'metre') and not part of mm/cm
    if 'mm' not in value_lower and 'cm' not in value_lower:
        if 'meter' in value_lower or 'metre' in value_lower:
            return num_value, 'meters'
        elif 'm' in value_lower:
            return num_value, 'meters'
    
    # Check for feet - check for 'ft', 'feet', or 'foot'
    if 'ft' in value_lower or 'feet' in value_lower or 'foot' in value_lower:
        return num_value, 'feet'
    
    # Check for inches - check for 'in' or 'inch' (but not part of other units)
    if ('in' in value_lower or 'inch' in value_lower) and 'mm' not in value_lower and 'cm' not in value_lower:
        return num_value, 'inches'
    
    # No unit detected - try to extract number anyway
    match = re.search(r'(\d+(?:\.\d+)?)', value_str)
    if match:
        num_value = float(match.group(1))
        return num_value, None
    
    return None, None


def validate_product_exists(base_product: str, available_models: List[str], 
                            price_calculator: Optional[PriceCalculator] = None,
                            filter_type: Optional[str] = None,
                            has_wd: Optional[bool] = None) -> Tuple[bool, Optional[str], bool, Optional[str]]:
    """
    Validate if a product exists in the available models and return normalized product info.
    Optionally validates filter if price_calculator is provided.
    
    Args:
        base_product: Base product name to validate (without WD, INS, or filter suffixes)
        available_models: List of available product models
        price_calculator: Optional PriceCalculator for filter validation
        filter_type: Optional filter type to validate
        has_wd: Optional pre-extracted WD flag (avoids redundant extraction)
        
    Returns:
        Tuple of (exists, normalized_product_name, has_wd_flag, error_message)
        If product doesn't exist, normalized_product_name will be None
        error_message will be None if validation passes, or contain error text if validation fails
    """
    if not available_models:
        return False, None, False, 'Price database not initialized'
    
    # Validate that base_product is not empty
    if base_product is None or not base_product.strip():
        return False, None, False, f'Invalid product name: "{base_product}"'
    
    # Extract WD flag if missing
    if has_wd is None:
        _, has_wd, _, _ = extract_product_flags_and_filter(base_product)
    
    # Validate filter if present and price_calculator is available
    if filter_type and price_calculator:
        if not validate_filter_exists(filter_type, price_calculator):
            return False, None, False, f'Filter "{filter_type}" not found in database'
    
    # Use find_matching_product to handle all matching logic (WD variants, base product, fuzzy matching)
    match_result = find_matching_product(base_product, available_models, has_wd)
    if match_result:
        matched_product, matched_has_wd = match_result
        return True, matched_product, matched_has_wd, None
    
    return False, None, False, f'Product "{base_product}" not found in database'

