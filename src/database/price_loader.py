"""
Price List Loader Module
Handles loading and parsing the SQLite price database for HRG and WSG products.
"""

import re
import sqlite3
from pathlib import Path


class PriceListLoader:
    """Loads and parses the price list from SQLite database"""
    
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
            SELECT anodized_multiplier, powder_coated_multiplier, other_paint_multiplier 
            FROM products 
            WHERE model = ? 
            LIMIT 1
        ''', (product,))
        
        result = cursor.fetchone()
        if not result:
            return []
        
        available_finishes = []
        anodized_multiplier, powder_coated_multiplier, other_paint_multiplier = result
        
        # Check if anodized aluminum is available (multiplier is not None)
        if anodized_multiplier is not None:
            available_finishes.append('Anodized Aluminum')
        
        # Check if white powder coated is available (multiplier is not None)
        if powder_coated_multiplier is not None:
            available_finishes.append('White Powder Coated')
        
        # Check if other paint is available (multiplier is not None)
        if other_paint_multiplier is not None:
            available_finishes.append('Other Paint')
        
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
    
    def get_price_for_default_table(self, product, finish, size, with_damper=False):
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
        
        # Get table_id and multipliers for the product (this should be fast with model index)
        cursor.execute('SELECT table_id, anodized_multiplier, powder_coated_multiplier, other_paint_multiplier FROM products WHERE model = ? LIMIT 1', (product,))
        table_result = cursor.fetchone()
        if not table_result:
            return 0
        
        table_id = table_result[0]
        anodized_multiplier = table_result[1]
        powder_coated_multiplier = table_result[2]
        other_paint_multiplier = table_result[3]
        
        # Get price multiplier for finish
        multiplier = 1.0
        if finish == 'Anodized Aluminum' and anodized_multiplier is not None:
            multiplier = anodized_multiplier
        elif finish == 'White Powder Coated' and powder_coated_multiplier is not None:
            multiplier = powder_coated_multiplier
        elif finish == 'Other Paint' and other_paint_multiplier is not None:
            multiplier = other_paint_multiplier
        
        # Query prices using the table_id and dimensions
        if with_damper:
            cursor.execute('''
                SELECT price_with_damper
                FROM prices
                WHERE table_id = ? AND width = ? AND height = ?
            ''', (table_id, width, height))
        else:
            cursor.execute('''
                SELECT normal_price
                FROM prices
                WHERE table_id = ? AND width = ? AND height = ?
            ''', (table_id, width, height))
        
        result = cursor.fetchone()
        if result and result[0] is not None:
            return int(result[0] * multiplier + 0.5)
        return 0
    
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
    
    def get_price_for_other_table(self, product, finish, diameter, with_damper=False):
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
        
        # Get table_id and multipliers for the product
        cursor.execute('SELECT table_id, anodized_multiplier, powder_coated_multiplier, other_paint_multiplier FROM products WHERE model = ? LIMIT 1', (product,))
        table_result = cursor.fetchone()
        if not table_result:
            return 0
        
        table_id = table_result[0]
        anodized_multiplier = table_result[1]
        powder_coated_multiplier = table_result[2]
        other_paint_multiplier = table_result[3]
        
        # Get price multiplier for finish
        multiplier = 1.0
        if finish == 'Anodized Aluminum' and anodized_multiplier is not None:
            multiplier = anodized_multiplier
        elif finish == 'White Powder Coated' and powder_coated_multiplier is not None:
            multiplier = powder_coated_multiplier
        elif finish == 'Other Paint' and other_paint_multiplier is not None:
            multiplier = other_paint_multiplier
        
        # Query prices using the table_id and diameter (height field for diameter-based products)
        if with_damper:
            cursor.execute('''
                SELECT price_with_damper
                FROM prices
                WHERE table_id = ? AND width IS NULL AND height = ?
            ''', (table_id, diameter))
        else:
            cursor.execute('''
                SELECT normal_price
                FROM prices
                WHERE table_id = ? AND width IS NULL AND height = ?
            ''', (table_id, diameter))
        
        result = cursor.fetchone()
        if result and result[0] is not None:
            return int(result[0] * multiplier + 0.5)
        return 0
    
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
        """Check if a product/finish combination has any non-null price_with_damper values"""
        conn = self._get_connection()
        if not conn:
            return False
        
        cursor = conn.cursor()
        
        # Get table_id for the product
        cursor.execute('SELECT table_id FROM products WHERE model = ? LIMIT 1', (product,))
        table_result = cursor.fetchone()
        if not table_result:
            return False
        
        table_id = table_result[0]
        
        # Check if there are any non-null price_with_damper values for this table
        cursor.execute('''
            SELECT COUNT(*) 
            FROM prices 
            WHERE table_id = ? AND price_with_damper IS NOT NULL
        ''', (table_id,))
        
        result = cursor.fetchone()
        return result[0] > 0 if result else False
    
    def __del__(self):
        """Clean up database connection when object is destroyed"""
        if self.conn:
            self.conn.close()

