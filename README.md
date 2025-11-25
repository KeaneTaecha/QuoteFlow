# Quotation System - Dynamic Product Management

A PyQt5-based quotation management system with dynamic product configuration using Excel Header sheet and SQLite database.

## Project Structure

```
QuoteFlow/
├── main.py                      # Entry point - run this!
├── quotation_ui.py              # UI components and main window
├── price_calculator.py              # Price calculator using SQLite database
├── getsql.py                    # Excel to SQLite converter
├── price_list_modified.xlsx     # Excel price data with Header sheet
├── prices.db                    # SQLite database (auto-generated)
├── requirements.txt             # Python dependencies
├── README.md                    # This file
├── HEADER_SHEET_GUIDE.md        # Header sheet documentation
├── CHANGES_SUMMARY.md           # Summary of changes
└── test_header_integration.py   # Integration test suite
```

## Features

- **Dynamic Product Configuration** via Excel Header sheet
- Load prices from SQLite database
- Support for multiple product models (1-HRG, 2-WSG, etc.)
- Multiple finishes (Anodized Aluminum, White Powder Coated)
- Automatic size rounding to nearest available size
- With Damper (WD) option support
- Quote management (add, remove, clear items)
- Save quotes to text files
- Professional quotation formatting
- Single or multiple models per price table

## Installation

1. **Install Python 3.8+** (if not already installed)

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Ensure Excel file with Header sheet is ready:**
   - The Excel file must contain a **Header** sheet
   - See `HEADER_SHEET_GUIDE.md` for details

## Initial Setup

Before running the application for the first time:

1. **Add Header sheet to your Excel file** (see Header Sheet Format below)

2. **Convert Excel to SQLite database:**
   ```bash
   python getsql.py
   ```
   This creates `prices.db` from your Excel file

3. **Verify the conversion:**
   ```bash
   python test_header_integration.py
   ```

## Header Sheet Format

The Excel file must include a **Header** sheet with these columns:

| Table id | Sheet | Model |
|----------|-------|-------|
| 1 | 1-HRG,2-WSG TB | 1-HRG, 2-WSG |

- **Table id**: Unique number for each price table
- **Sheet**: Name of the Excel sheet containing prices
- **Model**: Product model(s), comma-separated for multiple models

**Example Configurations:**

Single model:
```
Table id: 1
Sheet: 1-HRG TB
Model: 1-HRG
```
→ UI shows only "1-HRG"

Multiple models:
```
Table id: 1
Sheet: 1-HRG,2-WSG TB
Model: 1-HRG, 2-WSG
```
→ UI shows both "1-HRG" and "2-WSG"

See `HEADER_SHEET_GUIDE.md` for complete documentation.

## Running the Application

### As Python Script:
```bash
python main.py
```

### When to Regenerate Database:

Run `python getsql.py` again when:
- You add new products to the Header sheet
- You modify price data in Excel
- You add new price table sheets

## Usage

1. **Launch the application**
2. **Enter customer information** (name, quote number)
3. **Select product configuration:**
   - Product type (dynamically loaded from database)
   - Finish (Anodized Aluminum or White Powder Coated)
   - Width and height (in inches)
   - With Damper option
   - Quantity
4. **Click "Add to Quote"** to add items
5. **Save or print** your quote when ready

## How It Works

1. **Header Sheet** in Excel defines which products are available
2. **getsql.py** reads the Header sheet and converts Excel to SQLite
3. **Database** stores products and prices efficiently
4. **UI** dynamically loads available products from the database
5. **Price Calculator** calculates prices based on product and size selection

## Workflow

```
Excel (with Header) → getsql.py → prices.db → UI
```

1. Update Header sheet in Excel
2. Run `python getsql.py` to convert to database
3. Run `python main.py` to launch UI
4. UI shows products from database

## Notes

- The system automatically rounds dimensions to the nearest available size
- Prices are loaded from SQLite database at startup
- Products in the UI dropdown come from the Header sheet
- Multiple models can share the same price table
- Single model configuration: Only that model appears in UI

## Troubleshooting

**Error: Price database not found**
- Run: `python getsql.py` to create the database

**Error: No products found in database**
- Check that Header sheet exists in Excel file
- Verify Header sheet has data (not just column headers)
- Re-run: `python getsql.py`

**Error: Wrong products showing in UI**
- Check the Model column in Header sheet
- Ensure models are spelled correctly
- For multiple models, separate with comma: "1-HRG, 2-WSG"
- Re-run: `python getsql.py`

**Error: Failed to load price database**
- Delete `prices.db` and run `python getsql.py` again
- Check that Excel file exists and is not corrupted

**Module not found errors**
- Run: `pip install -r requirements.txt`

## Testing

Run the integration test suite:
```bash
python test_header_integration.py
```

This verifies:
- Header sheet reading
- Database structure
- Price calculator functionality
- UI integration

## Development

The code is modularized for easy maintenance:

- **getsql.py**: Excel to SQLite conversion, reads Header sheet
- **price_calculator.py**: Calculates prices using SQLite database
- **quotation_ui.py**: PyQt5 GUI components and user interactions
- **main.py**: Simple entry point that launches the application

### To Add New Products:

1. Edit the Header sheet in Excel
2. Run `python getsql.py`
3. Launch the UI - new products will appear

### To Modify Price Loading:

- Edit `price_calculator.py` for price calculation logic
- Edit `getsql.py` for database structure changes

### To Modify UI:

- Edit `quotation_ui.py` for interface changes

## Additional Documentation

- **HEADER_SHEET_GUIDE.md** - Complete guide to Header sheet format
- **CHANGES_SUMMARY.md** - Summary of all changes made
- **test_header_integration.py** - Test suite for verification

