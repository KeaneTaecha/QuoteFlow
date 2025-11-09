"""
Filter Price Utilities
Shared functions for calculating filter prices from the database.
"""

from typing import Optional
from utils.price_calculator import PriceCalculator


def get_filter_price(price_loader: PriceCalculator, filter_type: str, size_inches: float, 
                     width_inches: Optional[float] = None, height_inches: Optional[float] = None) -> Optional[float]:
    """
    Find filter product in database and get its price.
    
    Args:
        price_loader: PriceCalculator instance for database access
        filter_type: The filter type (e.g., "Nylon")
        size_inches: The size in inches (for diameter-based filters) or max dimension
        width_inches: Optional width in inches for dimension-based filters
        height_inches: Optional height in inches for dimension-based filters
        
    Returns:
        Filter price or None if not found
    """
    if not price_loader:
        return None
    
    # Get all available models
    all_models = price_loader.get_available_models()
    
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
    finishes = price_loader.get_available_finishes(matching_filter)
    if not finishes:
        return None
    
    # Use first available finish
    finish = finishes[0]
    
    # Check if filter is diameter-based (other table) or dimension-based
    is_other_table = price_loader.is_other_table(matching_filter)
    
    if is_other_table:
        # For diameter-based filters, get base price directly from database
        price = get_base_price_for_other_table(price_loader, matching_filter, size_inches)
        return price if price and price > 0 else None
    else:
        # For dimension-based filters, use actual width and height if provided
        if width_inches is not None and height_inches is not None:
            # Ensure width >= height (database convention)
            filter_width = max(width_inches, height_inches)
            filter_height = min(width_inches, height_inches)
            price = get_base_price_for_default_table(price_loader, matching_filter, filter_width, filter_height)
        else:
            # Fallback: use size_inches for both dimensions (square filter)
            price = get_base_price_for_default_table(price_loader, matching_filter, size_inches, size_inches)
        return price if price and price > 0 else None


def get_base_price_for_other_table(price_loader: PriceCalculator, product: str, diameter_inches: float) -> Optional[float]:
    """
    Get base price for other table product directly from database.
    
    Args:
        price_loader: PriceCalculator instance for database access
        product: Product model name
        diameter_inches: Diameter in inches
        
    Returns:
        Base price or None if not found
    """
    db = price_loader.db
    table_id = db.get_table_id(product)
    if table_id is None:
        return None
    
    diameter_int = int(diameter_inches)
    price_result = db.get_price_for_diameter(table_id, diameter_int)
    
    # If exact match not found, try to find closest
    if not price_result:
        # Find closest diameter >= given diameter
        closest_size = db.find_rounded_other_table_size(product, diameter_int)
        if closest_size:
            # Extract diameter from size string (e.g., "8\" diameter" -> 8)
            import re
            match = re.search(r'(\d+)', closest_size)
            if match:
                closest_diameter = int(match.group(1))
                price_result = db.get_price_for_diameter(table_id, closest_diameter)
    
    return price_result[0] if price_result else None


def get_base_price_for_default_table(price_loader: PriceCalculator, product: str, 
                                     width_inches: float, height_inches: float) -> Optional[float]:
    """
    Get base price for default table product directly from database.
    
    Args:
        price_loader: PriceCalculator instance for database access
        product: Product model name
        width_inches: Width in inches
        height_inches: Height in inches
        
    Returns:
        Base price or None if not found
    """
    db = price_loader.db
    table_id = db.get_table_id(product)
    if table_id is None:
        return None
    
    width_int = int(width_inches)
    height_int = int(height_inches)
    
    # Try exact match first
    price_result = db.get_price_for_dimensions(table_id, height_int, width_int)
    if price_result:
        return price_result[0]  # Return normal_price
    
    # If no exact match, find closest
    closest = db.find_closest_price_for_dimensions(table_id, height_int, width_int)
    if closest:
        return closest[2]  # Return normal_price from closest match
    
    return None

