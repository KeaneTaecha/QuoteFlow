"""
AUTOMATIC EXCEL TO SQLITE CONVERTER
====================================
This script automatically reads your entire Excel file and converts
ALL sheets to SQLite database. Just run it once!

Usage:
    python convert_excel_to_sqlite.py

Output:
    Creates 'prices.db' with all your price data
"""

import sqlite3
import openpyxl
from pathlib import Path
from datetime import datetime


class ExcelToSQLiteConverter:
    """Automatically convert Excel price list to SQLite database"""
    
    def __init__(self, excel_file, db_file='prices.db'):
        self.excel_file = excel_file
        self.db_file = db_file
        self.conn = None
        
        # Will read from Header sheet
        self.header_data = []
        
        # Statistics
        self.stats = {
            'total_sheets': 0,
            'processed_sheets': 0,
            'skipped_sheets': 0,
            'total_prices': 0,
            'total_products': 0,
            'errors': []
        }
    
    def create_database(self):
        """Create database structure"""
        print("Creating database structure...")
        
        self.conn = sqlite3.connect(self.db_file)
        cursor = self.conn.cursor()
        
        # Products table - now using table_id from Header sheet
        cursor.execute('''
            CREATE TABLE products (
                table_id INTEGER PRIMARY KEY NOT NULL,
                model TEXT NOT NULL,
                sheet_name TEXT NOT NULL,
                UNIQUE(table_id, model)
            )
        ''')
        
        # Prices table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS prices (
                price_id INTEGER PRIMARY KEY AUTOINCREMENT,
                table_id INTEGER NOT NULL,
                width INTEGER NOT NULL,
                height INTEGER NOT NULL,
                normal_price REAL,
                price_with_damper REAL,
                FOREIGN KEY (table_id) REFERENCES products(table_id) ON UPDATE CASCADE,
                UNIQUE(table_id, width, height)
            )
        ''')
        
        # Indexes for fast lookups
        cursor.execute('''
            CREATE INDEX idx_price_lookup 
            ON prices(table_id, width, height)
        ''')
        
        cursor.execute('''
            CREATE INDEX idx_model_lookup 
            ON products(model)
        ''')
        
        
        self.conn.commit()
        print("✓ Database structure created\n")
    
    def process_table(self, sheet, table_id):
        """Process the price table data from the sheet"""
        cursor = self.conn.cursor()
        
        # Find inch widths in row 8 (index 7) and process each column immediately
        # Inch width headers are in row 8, starting from column B (index 2)
        price_count = 0
        max_col = sheet.max_column if sheet.max_column <= 100 else 100
        max_row = sheet.max_row if sheet.max_row <= 500 else 500
        
        for col in range(2, max_col):
            try:
                cell_value = sheet.cell(8, col).value
                if cell_value and isinstance(cell_value, str) and '"' in str(cell_value):
                    # Extract inch width from string like "4"" -> 4
                    width_str = str(cell_value).replace('"', '').strip()
                    width_inches = int(width_str)
                    
                    # Process all height rows for this width column
                    for row in range(10, max_row, 2):  # Process inch height rows (10, 12, 14, 16, 18...)
                        try:
                            # Get height from column A (inch height row like "4"", "6"", "8"", etc.)
                            height_cell = sheet.cell(row, 1).value
                            
                            # Check if this is a valid inch height row (should contain string like "4"", "6"", "8"")
                            if height_cell and isinstance(height_cell, str) and '"' in str(height_cell):
                                # Extract height from string like "4"" -> 4 (inch unit)
                                try:
                                    height_str = str(height_cell).replace('"', '').strip()
                                    height = int(height_str)
                                    
                                    # Get normal price and damper price for this width
                                    try:
                                        # Normal price (current row - inch height row)
                                        normal_price_cell = sheet.cell(row, col).value
                                        normal_price = float(normal_price_cell) if normal_price_cell and isinstance(normal_price_cell, (int, float)) else None
                                        
                                        # Price with damper (next row - mm height row)
                                        damper_price_cell = sheet.cell(row + 1, col).value
                                        damper_price = float(damper_price_cell) if damper_price_cell and isinstance(damper_price_cell, (int, float)) else None
                                        
                                        # Insert if at least one price exists
                                        if normal_price is not None or damper_price is not None:
                                            cursor.execute('''
                                                INSERT OR REPLACE INTO prices 
                                                (table_id, width, height, normal_price, price_with_damper)
                                                VALUES (?, ?, ?, ?, ?)
                                            ''', (table_id, width_inches, height, normal_price, damper_price))
                                            price_count += 1
                                    except Exception as e:
                                        continue
                                except ValueError:
                                    break
                            else:
                                break
                        except:
                            break
                else:
                    break
            except:
                break
        
        # Print the end of table coordinates
        if price_count > 0:
            print(f"  ✓ Table ended at column {col}, row {row}")
        
        return price_count

    def read_header_sheet(self, wb):
        """Read the Header sheet to get table metadata"""
        print("Reading Header sheet...")
        
        if 'Header' not in wb.sheetnames:
            print("❌ Error: 'Header' sheet not found in Excel file!")
            return False
        
        header_sheet = wb['Header']
        self.header_data = []
        
        # Read header data (skip first row which contains column headers)
        for row in header_sheet.iter_rows(min_row=2, values_only=True):
            table_id, sheet_name, model = row[0], row[1], row[2]
            
            # Skip empty rows
            if table_id is None or sheet_name is None or model is None:
                continue
            
            # Parse models (can be comma-separated)
            models = [m.strip() for m in str(model).split(',')]
            
            self.header_data.append({
                'table_id': int(table_id),
                'sheet_name': str(sheet_name),
                'models': models
            })
        
        print(f"✓ Found {len(self.header_data)} table(s) in Header sheet\n")
        return True
    
    def insert_products(self, table_id, sheet_name, models):
        """Insert product records for a table"""
        cursor = self.conn.cursor()
        
        for model in models:
            cursor.execute('''
                INSERT OR IGNORE INTO products (table_id, model, sheet_name)
                VALUES (?, ?, ?)
            ''', (table_id, model, sheet_name))
            self.stats['total_products'] += 1
    
    def process_sheet(self, sheet, table_id, sheet_name, models):
        """Process a single Excel sheet"""
        try:
            # Insert products for this table
            self.insert_products(table_id, sheet_name, models)
            
            # Process the price table
            price_count = self.process_table(sheet, table_id)
            
            return price_count
            
        except Exception as e:
            self.stats['errors'].append(f"{sheet_name}: {str(e)}")
            return 0
    
    def convert(self):
        """Main conversion process"""
        print("="*70)
        print("AUTOMATIC EXCEL TO SQLITE CONVERTER")
        print("="*70)
        print(f"\nSource file: {self.excel_file}")
        print(f"Output database: {self.db_file}")
        print()
        
        # Check if Excel file exists
        if not Path(self.excel_file).exists():
            print(f"❌ Error: File '{self.excel_file}' not found!")
            return False
        
        # Create database
        self.create_database()
        
        # Load Excel file
        print(f"Loading Excel file...")
        try:
            wb = openpyxl.load_workbook(self.excel_file, read_only=True, data_only=True)
            print(f"✓ Loaded {len(wb.sheetnames)} sheets\n")
        except Exception as e:
            print(f"❌ Error loading Excel file: {e}")
            return False
        
        self.stats['total_sheets'] = len(wb.sheetnames)
        
        # Read Header sheet
        if not self.read_header_sheet(wb):
            return False
        
        if not self.header_data:
            print("❌ Error: No table information found in Header sheet!")
            return False
        
        # Process sheets based on Header information
        print("Processing sheets:")
        print("-" * 70)
        
        for header_entry in self.header_data:
            table_id = header_entry['table_id']
            sheet_name = header_entry['sheet_name']
            models = header_entry['models']
            
            if sheet_name not in wb.sheetnames:
                print(f"❌ Warning: Sheet '{sheet_name}' not found in Excel file!")
                self.stats['skipped_sheets'] += 1
                continue
            
            print(f"Processing: {sheet_name} (Table ID: {table_id}, Models: {', '.join(models)})...", end=" ")
            
            try:
                sheet = wb[sheet_name]
                price_count = self.process_sheet(sheet, table_id, sheet_name, models)
                
                if price_count > 0:
                    print(f"✓ {price_count} prices")
                    self.stats['processed_sheets'] += 1
                    self.stats['total_prices'] += price_count
                else:
                    print(f"⚠ No valid prices found")
                    self.stats['skipped_sheets'] += 1
                    
            except Exception as e:
                print(f"❌ Error: {str(e)}")
                self.stats['errors'].append(f"{sheet_name}: {str(e)}")
                self.stats['skipped_sheets'] += 1
        
        # Commit all changes
        self.conn.commit()
        self.conn.close()
        
        # Print summary
        self.print_summary()
        
        return True
    
    def print_summary(self):
        """Print conversion summary"""
        print("\n" + "="*70)
        print("CONVERSION COMPLETE!")
        print("="*70)
        print(f"\nTotal sheets in Excel:     {self.stats['total_sheets']}")
        print(f"Sheets processed:          {self.stats['processed_sheets']}")
        print(f"Sheets skipped:            {self.stats['skipped_sheets']}")
        print(f"Total products created:    {self.stats['total_products']}")
        print(f"Total prices inserted:     {self.stats['total_prices']:,}")
        
        if self.stats['errors']:
            print(f"\nErrors encountered:        {len(self.stats['errors'])}")
            for error in self.stats['errors'][:5]:
                print(f"  - {error}")
        
        print(f"\n✓ Database saved to: {self.db_file}")
        print(f"✓ File size: {Path(self.db_file).stat().st_size / 1024:.1f} KB")
        print("\nYou can now use this database in your PyQt application!")


# =============================================================================
# MAIN PROGRAM - RUN THIS!
# =============================================================================

if __name__ == "__main__":
    # Configuration
    EXCEL_FILE = 'price_list_modified.xlsx'  # Your Excel file name
    DATABASE_FILE = 'prices.db'               # Output database name
    
    print("\n")
    print("╔════════════════════════════════════════════════════════════════════╗")
    print("║         AUTOMATIC EXCEL TO SQLITE CONVERTER                       ║")
    print("║         Easy conversion - Just run this script!                   ║")
    print("╚════════════════════════════════════════════════════════════════════╝")
    print("\n")
    
    # Create converter
    converter = ExcelToSQLiteConverter(EXCEL_FILE, DATABASE_FILE)
    
    # Run conversion
    success = converter.convert()
    
    if success:
        print("\n✓ Conversion completed successfully!")
    else:
        print("\n❌ Conversion failed. Please check the errors above.")