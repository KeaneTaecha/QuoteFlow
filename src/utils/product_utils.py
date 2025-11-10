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
    
    Note: Uses division by 25 (not 25.4) to match existing codebase behavior.
    
    Args:
        value_mm: Value in millimeters
        
    Returns:
        Value in inches
    """
    return value_mm / 25


def convert_dimension_to_inches(value: float, unit: str) -> float:
    """
    Convert a dimension value to inches based on the unit.
    
    Args:
        value: The dimension value
        unit: Unit string ('Millimeters', 'millimeters', 'mm', 'Inches', 'inches', etc.)
        
    Returns:
        Value in inches
    """
    unit_lower = str(unit).lower()
    if 'mm' in unit_lower or 'millimeter' in unit_lower:
        return convert_mm_to_inches(value)
    else:
        # Already in inches
        return value


def find_matching_product(product: str, available_models: List[str], 
                          base_product: Optional[str] = None, 
                          has_wd: Optional[bool] = None) -> Optional[Tuple[str, bool]]:
    """
    Find a matching product in the available models list using exact match only.
    
    Args:
        product: Product name to find
        available_models: List of available product models
        base_product: Optional pre-extracted base product name (unused, kept for compatibility)
        has_wd: Optional pre-extracted WD flag (unused, kept for compatibility)
        
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


def validate_filter_exists(filter_type: str, price_loader: PriceCalculator) -> bool:
    """
    Validate if a filter type exists in the database.
    
    Args:
        filter_type: Filter type to validate (e.g., "Nylon")
        price_loader: PriceCalculator instance for database access
        
    Returns:
        True if filter exists, False otherwise
    """
    if not filter_type or not price_loader:
        return False
    
    all_models = price_loader.get_available_models()
    filter_type_lower = filter_type.lower()
    
    # Search for filter product that matches the filter type
    for model in all_models:
        if filter_type_lower in model.lower():
            return True
    
    return False


def get_product_type_flags(price_loader: PriceCalculator, product: str) -> Tuple[bool, bool, bool]:
    """
    Get product type flags (has_no_dimensions, has_price_per_foot, is_other_table).
    Consolidates the logic for determining product type characteristics.
    
    Args:
        price_loader: PriceCalculator instance
        product: Product model name
        
    Returns:
        Tuple of (has_no_dimensions, has_price_per_foot, is_other_table)
    """
    has_no_dimensions = price_loader.has_no_dimensions(product)
    has_price_per_foot = price_loader.has_price_per_foot(product)
    
    # Determine is_other_table based on product characteristics
    if has_no_dimensions:
        is_other_table = price_loader.is_other_table(product)
    else:
        # Only check if it's NOT a price_per_foot product
        is_other_table = price_loader.is_other_table(product) if not has_price_per_foot else False
    
    return has_no_dimensions, has_price_per_foot, is_other_table


def validate_product_exists(product: str, available_models: List[str], 
                            price_loader: Optional[PriceCalculator] = None,
                            filter_type: Optional[str] = None,
                            base_product: Optional[str] = None,
                            has_wd: Optional[bool] = None) -> Tuple[bool, Optional[str], bool, Optional[str]]:
    """
    Validate if a product exists in the available models and return normalized product info.
    Optionally validates filter if price_loader is provided.
    
    Args:
        product: Product name to validate
        available_models: List of available product models
        price_loader: Optional PriceCalculator for filter validation
        filter_type: Optional filter type to validate
        base_product: Optional pre-extracted base product name (avoids redundant extraction)
        has_wd: Optional pre-extracted WD flag (avoids redundant extraction)
        
    Returns:
        Tuple of (exists, normalized_product_name, has_wd_flag, error_message)
        If product doesn't exist, normalized_product_name will be None
        error_message will be None if validation passes, or contain error text if validation fails
    """
    if not available_models:
        return False, None, False, 'Price database not initialized'
    
    # Extract base product and WD flag if missing (filter_type can be None, so only extract if base_product/has_wd missing)
    if base_product is None or has_wd is None:
        extracted_base, extracted_wd, _, extracted_filter = extract_product_flags_and_filter(product)
        
        # Validate extraction results
        if extracted_base is None or not extracted_base.strip():
            return False, None, False, f'Failed to extract product name from "{product}"'
        
        base_product = base_product if base_product is not None else extracted_base
        has_wd = has_wd if has_wd is not None else extracted_wd
        # Only use extracted filter if filter_type was not provided (use provided filter_type even if None)
        if filter_type is None:
            filter_to_validate = extracted_filter
        else:
            filter_to_validate = filter_type
    else:
        # base_product and has_wd provided, use them directly
        filter_to_validate = filter_type  # Can be None, that's valid
        # Validate that base_product is not empty
        if base_product is None or not base_product.strip():
            return False, None, False, f'Invalid product name: "{product}"'
    
    # Validate filter if present and price_loader is available
    if filter_to_validate and price_loader:
        if not validate_filter_exists(filter_to_validate, price_loader):
            return False, None, False, f'Filter "{filter_to_validate}" not found in database'
    
    # Use find_matching_product to handle all matching logic (WD variants, base product, fuzzy matching)
    match_result = find_matching_product(product, available_models, base_product, has_wd)
    if match_result:
        matched_product, matched_has_wd = match_result
        return True, matched_product, matched_has_wd, None
    
    return False, None, False, f'Product "{product}" not found in database'

