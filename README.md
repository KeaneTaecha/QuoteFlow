# QuoteFlow - Quotation Management System

QuoteFlow is a desktop application for managing product quotations with automated pricing calculations. Built with PyQt5, it provides an intuitive interface for creating, managing, and exporting professional quotations based on a comprehensive product database.

## Features

- **Product Database Management**: Import and manage product pricing from Excel files
- **Automated Price Calculation**: Calculate prices based on dimensions, finishes, and modifiers
- **Excel Import/Export**: Import items from Excel and export professional quotations
- **Filter System**: Support for various filter types and damper configurations
- **Finish Multipliers**: Automatic price adjustments for Anodized, Powder Coated, and No Finish options
- **Quotation Management**: Create, edit, and save multiple quotations
- **Professional Output**: Export quotations to formatted Excel files

## System Requirements

- Python 3.7 or higher
- Windows / macOS / Linux

## Installation

### 1. Clone or Download the Repository

```bash
git clone <repository-url>
cd QuoteFlow
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

Required packages:
- PyQt5 >= 5.15.10
- openpyxl >= 3.1.2
- Pillow >= 9.0.0
- pyinstaller >= 6.3.0 (for building executables)

### 3. Prepare Price Database

Before running the application, you need to create the price database from your Excel price list:

```bash
python src/excel_to_sql/getsql.py
```

This will convert your Excel price list into a SQLite database (`prices.db`). See [EXCEL_PRICE_TEMPLATE_README.md](/docs/EXCEL_PRICE_TEMPLATE_README.md) for details on the Excel format requirements.

## Usage

### Running the Application

```bash
python src/main.py
```

Or simply:

```bash
python -m src.main
```

### Basic Workflow

1. **Select Product Model**: Choose from available product models in the dropdown
2. **Configure Dimensions**: Enter width and height in inches or millimeters
3. **Select Finish**: Choose Anodized Aluminum, Powder Coated, No Finish, or Special Color
4. **Set Quantity**: Specify the number of units
5. **Add Filters** (Optional): Configure filter options with damper settings
6. **Add to Quotation**: Click "Add Item" to add to the current quotation
7. **Export**: Generate professional Excel quotation using "Export to Excel"

### Importing Items from Excel

QuoteFlow supports bulk import of items from Excel files:

1. Click "Import from Excel" button
2. Select your Excel file containing items (see format requirements below)
3. Review imported items and any warnings
4. Items are automatically added to your current quotation

For Excel import format specifications, see [EXCEL_QUOTATION_IMPORT_README.md](/docs/EXCEL_QUOTATION_IMPORT_README.md)

## Project Structure

```
QuoteFlow/
├── assets/                     # Application icons
├── data/                       # Sample price lists and Excel templates
│   ├── Price List Update FEB 2024_Modified.xlsx
│   ├── Price List Update NOV 2025.xls
│   └── quotation_template.xlsx
├── docs/                       # Documentation
│   ├── EXCEL_PRICE_TEMPLATE_README.md
│   └── EXCEL_QUOTATION_IMPORT_README.md
├── src/
│   ├── main.py                 # Application entry point
│   ├── excel_to_sql/           # Excel to SQLite converter
│   │   ├── getsql.py
│   │   ├── excel_utils.py
│   │   ├── table_models.py
│   │   └── handlers/
│   │       ├── default_handler.py
│   │       ├── header_handler.py
│   │       └── other_handler.py
│   ├── ui/
│   │   └── quotation_ui.py     # Main PyQt5 UI
│   ├── utils/
│   │   ├── equation_parser.py
│   │   ├── excel_exporter.py
│   │   ├── excel_importer.py
│   │   ├── filter_utils.py
│   │   ├── price_calculator.py
│   │   ├── product_utils.py
│   │   ├── quote_utils.py
│   │   └── sql_loader.py
│   └── Quotation_25-12023.xlsx # Sample quotation export
├── scripts/
│   └── rename_tb_modifier_column.py
├── create_icons.py             # Icon generation helper
├── requirements.txt            # Python dependencies
├── QuoteFlow.spec              # PyInstaller build specification
├── prices.db                   # Generated SQLite database (created by getsql.py)
└── README.md
```

## Excel Templates

QuoteFlow uses two types of Excel templates:

### 1. Price List Template

Used for importing product pricing data into the database. This file contains:
- Multiple sheets with different product lines
- Header sheet with product metadata
- Standard and Other table formats for pricing matrices

See [EXCEL_PRICE_TEMPLATE_README.md](/docs/EXCEL_PRICE_TEMPLATE_README.md) for detailed format specifications

### 2. Quotation Import Template

Used for importing multiple items into a quotation at once:
- Flexible column ordering
- Support for dimensions, finishes, and quantities
- Title rows for organization

See [EXCEL_QUOTATION_IMPORT_README.md](/docs/EXCEL_QUOTATION_IMPORT_README.md) for detailed format specifications

## Database Structure

QuoteFlow uses SQLite for storing product and pricing data:

### Tables

- **products**: Product models with finish multipliers and modifiers
- **prices**: Price matrices based on dimensions
- **row_multipliers**: Width-specific multipliers for height exceeded scenarios
- **column_multipliers**: Height-specific multipliers for width exceeded scenarios

### Price Calculation Logic

1. Base price lookup from dimensions (width × height)
2. Apply base modifier equation if defined
3. Apply finish multiplier (Anodized Aluminum/Powder Coated/No Finish/Special Color)
4. Add filter price if selected
5. Apply quantity discount if applicable
6. Calculate final total

## Building Executable

To create a standalone executable:

```bash
pyinstaller QuoteFlow.spec
```

The executable will be created in the `dist/` directory.

## Features in Detail

### Product Management
- Automatic product validation against database
- Support for multiple product types (standard, filter, damper)
- Slot-based products with special pricing rules

### Dimension Handling
- Support for both inches (default) and millimeters
- Automatic unit conversion
- Precision matching with database prices

### Filter System
- Multiple filter types: 2-inch, 4-inch, Activated Carbon, Poly
- Damper configurations (with/without)
- Automatic filter price calculation

### Finish Options
- **Anodized Aluminum**: Default multiplier typically 1.0
- **Powder Coated**: Custom multiplier per product
- **No Finish**: Raw aluminum pricing

### Discount System
- Per-item discount percentages
- Automatic price recalculation
- Reflected in quotation totals

## Troubleshooting

### Database Issues

**Problem**: "No price found for product X"
- Ensure your price list Excel is properly formatted
- Run `getsql.py` to regenerate the database
- Check that dimensions match available prices

**Problem**: "Product not found in database"
- Verify product model exists in Header sheet
- Check spelling and format (case-sensitive)
- Ensure database was generated from latest Excel

### Excel Import Issues

**Problem**: Import fails or skips items
- Check Excel format matches specifications
- Ensure required columns (Model) are present
- Review error messages in import dialog

**Problem**: Prices not calculating correctly
- Verify dimensions are in correct unit
- Check finish selection matches database multipliers
- Ensure no blank or invalid data in Excel

### Performance Issues

**Problem**: Application slow to start
- Database may be too large - optimize price list
- Check for corrupted database file
- Reduce number of products if possible

## Development

### Modifying Price Calculations

Price calculation logic is in `src/utils/price_calculator.py`:
- `calculate_product_price()`: Main pricing logic
- Modify modifier equations in database Header sheet
- Adjust multipliers per product as needed

### Adding New Table Types

To support new Excel table formats:
1. Create new handler in `src/excel_to_sql/handlers/`
2. Implement table detection and extraction logic
3. Register handler in `getsql.py`

### UI Customization

Main UI is in `src/ui/quotation_ui.py`:
- PyQt5 widgets for all controls
- Modify layouts and styling as needed
- Font scaling system for accessibility



---

**Important**: Before first use, generate the price database using `python src/excel_to_sql/getsql.py` with your properly formatted Excel price list.
