"""
Database Access Module
Handles all database operations for the price database.
Separated from price calculation logic for better code organization.
"""

import sqlite3
from pathlib import Path
from typing import Optional, List, Tuple


class PriceDatabase:
    """Handles all database operations for price queries"""
    
    def __init__(self, db_path='../prices.db'):
        self.db_path = db_path
        self.conn = None
        self._check_database()
    
    def _check_database(self):
        """Check if database exists and is accessible"""
        if not Path(self.db_path).exists():
            print(f"❌ Error: Database file '{self.db_path}' not found!")
            return False
        return True
    
    def get_connection(self):
        """Get database connection, creating one if needed"""
        if self.conn is None:
            if not self._check_database():
                return None
            self.conn = sqlite3.connect(self.db_path)
        return self.conn
    
    def close(self):
        """Close database connection"""
        if self.conn:
            self.conn.close()
            self.conn = None
    
    # Product queries
    def get_available_models(self) -> List[str]:
        """Get list of available product models"""
        conn = self.get_connection()
        if not conn:
            return []
        
        cursor = conn.cursor()
        cursor.execute('SELECT DISTINCT model FROM products ORDER BY model')
        return [row[0] for row in cursor.fetchall()]
    
    def get_available_finishes(self, product: str) -> List[str]:
        """Get list of available finish options for a specific product"""
        conn = self.get_connection()
        if not conn:
            return []
        
        cursor = conn.cursor()
        cursor.execute('''
            SELECT anodized_multiplier, powder_coated_multiplier, no_finish_multiplier 
            FROM products 
            WHERE model = ? 
            LIMIT 1
        ''', (product,))
        
        result = cursor.fetchone()
        if not result:
            return []
        
        available_finishes = []
        anodized_multiplier, powder_coated_multiplier, no_finish_multiplier = result
        
        # Check if no finish is available (multiplier is not None)
        if no_finish_multiplier is not None:
            available_finishes.append('No Finish')
        
        # Check if anodized aluminum is available (multiplier is not None)
        if anodized_multiplier is not None:
            available_finishes.append('Anodized Aluminum')
        
        # Check if powder coated is available (multiplier is not None)
        if powder_coated_multiplier is not None:
            available_finishes.append('Powder Coated')
        
        # Always add Special Color option
        available_finishes.append('Special Color')
        
        return available_finishes
    
    def check_product_condition(self, product: str, condition_sql: str) -> bool:
        """Helper method to check product conditions in database"""
        # Validate product input
        if not product or not str(product).strip():
            return False
        
        conn = self.get_connection()
        if not conn:
            return False
        
        try:
            cursor = conn.cursor()
            cursor.execute(f'''
                SELECT COUNT(*) 
                FROM products p
                JOIN prices pr ON p.table_id = pr.table_id
                WHERE p.model = ? AND {condition_sql}
            ''', (product,))
            
            result = cursor.fetchone()
            return result[0] > 0 if result else False
        except Exception:
            # Return False on any database error
            return False
    
    def is_other_table(self, product: str) -> bool:
        """Check if a product uses other table format (diameter-based) instead of width/height"""
        return self.check_product_condition(product, 'pr.width IS NULL')
    
    def has_price_per_foot(self, product: str) -> bool:
        """Check if a product has price_per_foot pricing"""
        return self.check_product_condition(product, 'pr.price_per_foot IS NOT NULL')
    
    def has_no_dimensions(self, product: str) -> bool:
        """Check if a product has no height and width (both are NULL)"""
        return self.check_product_condition(product, 'pr.height IS NULL AND pr.width IS NULL')
    
    def has_damper_option(self, product: str) -> bool:
        """Check if a product has a non-null WD multiplier in the header sheet"""
        conn = self.get_connection()
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
    
    def get_price_id_for_no_dimensions(self, product: str) -> Optional[int]:
        """Get price_id for a product with no height/width by calculating from product_id difference
        
        Args:
            product: Product model name
            
        Returns:
            price_id or None if not found
        """
        conn = self.get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        
        # Get current product_id and table_id for the product
        cursor.execute('SELECT product_id, table_id FROM products WHERE model = ? LIMIT 1', (product,))
        product_result = cursor.fetchone()
        if not product_result:
            return None
        
        current_product_id = product_result[0]
        table_id = product_result[1]
        
        # Get the first product_id that uses the table_id
        cursor.execute('SELECT MIN(product_id) FROM products WHERE table_id = ?', (table_id,))
        first_product_result = cursor.fetchone()
        if not first_product_result or first_product_result[0] is None:
            return None
        
        first_product_id = first_product_result[0]
        
        # Calculate difference
        difference = current_product_id - first_product_id
        
        # Get the first price_id that uses the table_id (where height IS NULL and width IS NULL)
        cursor.execute('''
            SELECT MIN(price_id)
            FROM prices
            WHERE table_id = ? AND height IS NULL AND width IS NULL
        ''', (table_id,))
        
        first_price_result = cursor.fetchone()
        if not first_price_result or first_price_result[0] is None:
            return None
        
        first_price_id = first_price_result[0]
        
        # Calculate target price_id
        target_price_id = first_price_id + difference
        
        return target_price_id
    
    # Product data queries
    def get_product_data(self, product: str) -> Optional[Tuple]:
        """Get product data (table_id, base_modifier, anodized_multiplier, powder_coated_multiplier, no_finish_multiplier, wd_multiplier)
        
        Returns:
            Tuple of (table_id, base_modifier, anodized_multiplier, powder_coated_multiplier, no_finish_multiplier, wd_multiplier) or None
        """
        conn = self.get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        cursor.execute('SELECT table_id, base_modifier, anodized_multiplier, powder_coated_multiplier, no_finish_multiplier, wd_multiplier FROM products WHERE model = ? LIMIT 1', (product,))
        return cursor.fetchone()
    
    def get_product_multipliers(self, product: str) -> Optional[Tuple]:
        """Get product multipliers (anodized_multiplier, powder_coated_multiplier, no_finish_multiplier)
        
        Returns:
            Tuple of (anodized_multiplier, powder_coated_multiplier, no_finish_multiplier) or None
        """
        conn = self.get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        cursor.execute('SELECT anodized_multiplier, powder_coated_multiplier, no_finish_multiplier FROM products WHERE model = ? LIMIT 1', (product,))
        return cursor.fetchone()
    
    def get_table_id(self, product: str) -> Optional[int]:
        """Get table_id for a product"""
        conn = self.get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        cursor.execute('SELECT table_id FROM products WHERE model = ? LIMIT 1', (product,))
        result = cursor.fetchone()
        return result[0] if result else None
    
    # Price queries
    def get_price_for_dimensions(self, table_id: int, height: float, width: float) -> Optional[Tuple[float, float]]:
        """Get tb_price and wd_price for given dimensions
        
        Returns:
            Tuple of (tb_price, wd_price) or None if not found
        """
        conn = self.get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        cursor.execute('''
            SELECT normal_price, price_with_damper
            FROM prices
            WHERE table_id = ? AND height = ? AND width = ?
        ''', (table_id, height, width))
        
        result = cursor.fetchone()
        if not result or (result[0] is None and result[1] is None):
            return None
        # Return as (tb_price, wd_price) for consistency
        tb_price = result[0] if result[0] is not None else 0
        wd_price = result[1] if result[1] is not None else 0
        return tb_price, wd_price
    
    def get_price_for_diameter(self, table_id: int, diameter: float) -> Optional[Tuple[float, float]]:
        """Get tb_price and wd_price for given diameter (other table)
        
        Returns:
            Tuple of (tb_price, wd_price) or None if not found
        """
        conn = self.get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        cursor.execute('''
            SELECT normal_price, price_with_damper
            FROM prices
            WHERE table_id = ? AND height = ? AND width IS NULL
        ''', (table_id, diameter))
        
        result = cursor.fetchone()
        if not result or (result[0] is None and result[1] is None):
            return None
        # Return as (tb_price, wd_price) for consistency
        tb_price = result[0] if result[0] is not None else 0
        wd_price = result[1] if result[1] is not None else 0
        return tb_price, wd_price
    
    def get_price_per_foot(self, table_id: int, width: float, price_id: Optional[int] = None) -> Optional[float]:
        """Get price_per_foot for given table_id and width, or by price_id
        
        Always uses height column to get size since price_per_foot products store size in height.
        
        Returns:
            price_per_foot value or None
        """
        conn = self.get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        
        if price_id is not None:
            cursor.execute('''
                SELECT price_per_foot
                FROM prices
                WHERE price_id = ? AND price_per_foot IS NOT NULL
            ''', (price_id,))
        else:
            cursor.execute('''
                SELECT price_per_foot
                FROM prices
                WHERE table_id = ? AND height = ? AND price_per_foot IS NOT NULL
                ORDER BY height
                LIMIT 1
            ''', (table_id, width))
        
        result = cursor.fetchone()
        return result[0] if result else None
    
    def get_price_by_id(self, price_id: int) -> Optional[Tuple]:
        """Get price data by price_id
        
        Returns:
            Tuple of price data or None
        """
        conn = self.get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM prices WHERE price_id = ?', (price_id,))
        return cursor.fetchone()
    
    def get_diameter_by_price_id(self, price_id: int) -> Optional[float]:
        """Get diameter (height) for a price_id in other table format"""
        conn = self.get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        cursor.execute('''
            SELECT height
            FROM prices
            WHERE price_id = ? AND width IS NULL
        ''', (price_id,))
        
        result = cursor.fetchone()
        return result[0] if result else None
    
    # Size lookup queries
    def find_rounded_price_per_foot_width(self, product: str, width: float) -> Optional[float]:
        """Find the exact match first, then the next available width that is >= the given width for price_per_foot products
        
        Always searches in height column since price_per_foot products store size in height.
        """
        table_id = self.get_table_id(product)
        if table_id is None:
            return None
        
        conn = self.get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT height,
                   CASE 
                       WHEN height = ? THEN 0
                       ELSE height - ?
                   END as priority
            FROM prices
            WHERE table_id = ? AND price_per_foot IS NOT NULL AND height >= ?
            ORDER BY priority, height
            LIMIT 1
        ''', (width, width, table_id, width))
        
        result = cursor.fetchone()
        return result[0] if result else None
    
    def find_rounded_default_table_size(self, product: str, width: float, height: float) -> Optional[str]:
        """Find the exact match first, then the next available size that is >= the given width and height"""
        conn = self.get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        
        # Single query: prioritize exact match, then closest >= match
        cursor.execute('''
            SELECT pr.height, pr.width,
                   CASE 
                       WHEN pr.height = ? AND pr.width = ? THEN 0
                       ELSE (pr.height - ?) + (pr.width - ?)
                   END as priority
            FROM products p
            JOIN prices pr ON p.table_id = pr.table_id
            WHERE p.model = ? AND pr.height >= ? AND pr.width >= ?
            ORDER BY priority, pr.height, pr.width
            LIMIT 1
        ''', (height, width, height, width, product, height, width))
        
        result = cursor.fetchone()
        if result:
            # Return format: width x height (height after width)
            # Preserve decimal values if present
            width_val = result[1]
            height_val = result[0]
            # Format as integer if whole number, otherwise preserve decimals
            width_str = f'{int(width_val)}"' if width_val == int(width_val) else f'{width_val}"'
            height_str = f'{int(height_val)}"' if height_val == int(height_val) else f'{height_val}"'
            return f'{width_str} x {height_str}'
        
        return None
    
    def find_rounded_other_table_size(self, product: str, diameter: float, price_id: Optional[int] = None) -> Optional[str]:
        """Find the exact match first, then the next available diameter that is >= the given diameter for other table products
        
        Args:
            product: Product model name
            diameter: Diameter in inches (not used if price_id is provided)
            price_id: Optional price_id to use directly (for has_no_dimensions case)
            
        Returns:
            Size string (e.g., "8\" diameter") or None if not found
        """
        conn = self.get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        
        # If price_id is provided, use it directly (for has_no_dimensions case)
        if price_id is not None:
            diameter_result = self.get_diameter_by_price_id(price_id)
            if diameter_result:
                # Format as integer if whole number, otherwise preserve decimals
                diameter_str = f'{int(diameter_result)}"' if diameter_result == int(diameter_result) else f'{diameter_result}"'
                return f'{diameter_str} diameter'
            return None
        
        # Single query: prioritize exact match, then closest >= match
        cursor.execute('''
            SELECT pr.height,
                   CASE 
                       WHEN pr.height = ? THEN 0
                       ELSE pr.height - ?
                   END as priority
            FROM products p
            JOIN prices pr ON p.table_id = pr.table_id
            WHERE p.model = ? AND pr.height >= ? AND pr.width IS NULL
            ORDER BY priority, pr.height
            LIMIT 1
        ''', (diameter, diameter, product, diameter))
        
        result = cursor.fetchone()
        if result:
            diameter_val = result[0]
            # Format as integer if whole number, otherwise preserve decimals
            diameter_str = f'{int(diameter_val)}"' if diameter_val == int(diameter_val) else f'{diameter_val}"'
            return f'{diameter_str} diameter'
        
        return None
    
    # Multiplier queries
    def get_exceeded_dimension_multiplier(self, table_id: int, width: int, height: int, with_damper: bool = False) -> Optional[float]:
        """Get the appropriate multiplier when dimensions exceed table limits"""
        conn = self.get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        
        # Get the maximum dimensions available in the table
        cursor.execute('''
            SELECT MAX(height), MAX(width)
            FROM prices
            WHERE table_id = ?
        ''', (table_id,))
        
        max_dims = cursor.fetchone()
        if not max_dims or not max_dims[0] or not max_dims[1]:
            return None
        
        max_height, max_width = max_dims
        
        # Check if both dimensions exceed limits - use fallback pricing
        if width > max_width and height > max_height:
            # Use the highest available multipliers as fallback
            # Try to get the highest row multiplier (for height exceeded)
            cursor.execute('''
                SELECT height_exceeded_multiplier, height_exceeded_multiplier_wd
                FROM row_multipliers
                WHERE table_id = ?
                ORDER BY width DESC
                LIMIT 1
            ''', (table_id,))
            row_result = cursor.fetchone()
            
            # Try to get the highest column multiplier (for width exceeded)
            cursor.execute('''
                SELECT width_exceeded_multiplier, width_exceeded_multiplier_wd
                FROM column_multipliers
                WHERE table_id = ?
                ORDER BY height DESC
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
        
        # Check if only width exceeds limit - use specific height column multiplier
        elif width > max_width:
            # Find the closest height column that has a width exceeded multiplier
            # Use the actual height value to look up in column_multipliers
            cursor.execute('''
                SELECT width_exceeded_multiplier, width_exceeded_multiplier_wd
                FROM column_multipliers
                WHERE table_id = ? AND height <= ?
                ORDER BY height DESC
                LIMIT 1
            ''', (table_id, height))
            result = cursor.fetchone()
            if result:
                # Return appropriate multiplier based with_damper flag
                if with_damper and result[1] is not None:
                    return result[1]  # WD multiplier
                elif not with_damper and result[0] is not None:
                    return result[0]  # Regular multiplier
            return None
        
        # Check if only height exceeds limit - use specific width row multiplier
        elif height > max_height:
            # Find the closest width row that has a height exceeded multiplier
            # Use the actual width value to look up in row_multipliers
            cursor.execute('''
                SELECT height_exceeded_multiplier, height_exceeded_multiplier_wd
                FROM row_multipliers
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
    
    def get_max_dimensions(self, table_id: int) -> Optional[Tuple[float, float]]:
        """Get maximum height and width for a table_id"""
        conn = self.get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        cursor.execute('''
            SELECT MAX(height), MAX(width)
            FROM prices
            WHERE table_id = ?
        ''', (table_id,))
        
        result = cursor.fetchone()
        if not result or not result[0] or not result[1]:
            return None
        return (result[0], result[1])
    
    def find_closest_price_for_dimensions(self, table_id: int, height: float, width: float) -> Optional[Tuple[float, float, float, float]]:
        """Find closest available dimensions and prices
        
        Returns:
            Tuple of (height, width, tb_price, wd_price) or None if not found
        """
        conn = self.get_connection()
        if not conn:
            return None
        
        cursor = conn.cursor()
        cursor.execute('''
            SELECT height, width, normal_price, price_with_damper
            FROM prices
            WHERE table_id = ? AND height <= ? AND width <= ?
            ORDER BY height DESC, width DESC
            LIMIT 1
        ''', (table_id, height, width))
        
        result = cursor.fetchone()
        if result:
            # Return as (height, width, tb_price, wd_price) for consistency
            height = result[0]
            width = result[1]
            tb_price = result[2] if result[2] is not None else 0
            wd_price = result[3] if result[3] is not None else 0
            return height, width, tb_price, wd_price
        return None
    
    def __del__(self):
        """Clean up database connection when object is destroyed"""
        self.close()

