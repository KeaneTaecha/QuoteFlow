"""
Filter Price Utilities
Shared functions for calculating filter prices from the database.
"""

from typing import Optional
from utils.price_calculator import PriceCalculator


def get_filter_price(price_loader: PriceCalculator, filter_type: str, max_dimension: float,
                     width_inches: float, height_inches: float) -> Optional[float]:
    """
    Find filter product in database and get its price for default table products only.
    Uses get_price_for_default_table to handle exceeded dimensions correctly.
    
    Args:
        price_loader: PriceCalculator instance for database access
        filter_type: The filter type (e.g., "Nylon")
        max_dimension: The maximum dimension (max of width and height) in inches
        width_inches: Width in inches for dimension-based filters
        height_inches: Height in inches for dimension-based filters
        
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
    
    # Filters only apply to default table products (with width and height)
    # Ensure width >= height (database convention)
    filter_width = max(width_inches, height_inches)
    filter_height = min(width_inches, height_inches)
    
    # Build size string in the format expected by get_price_for_default_table
    size_str = f'{int(filter_width)}" x {int(filter_height)}"'
    
    try:
        # Use get_price_for_default_table which handles exceeded dimensions correctly
        # Pass finish=None and with_damper=False to get base price with modifiers applied
        price, _ = price_loader.get_price_for_default_table(matching_filter, None, size_str, with_damper=False)
        return price if price and price > 0 else None
    except Exception:
        # If price calculation fails, return None
        return None

