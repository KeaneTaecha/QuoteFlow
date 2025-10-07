# Quotation System - HRG & WSG Products

A PyQt5-based quotation management system for HRG and WSG products with Excel price list integration.

## Project Structure

```
Komfortflow/
├── main.py                      # Entry point - run this!
├── quotation_ui.py              # UI components and main window
├── price_loader.py              # Price list Excel parser
├── price_list_modified.xlsx     # Price data (required)
├── requirements.txt             # Python dependencies
├── build_exe.py                 # Build script for .exe
└── README.md                    # This file
```

## Features

- Load prices from Excel spreadsheet
- Support for HRG and WSG products
- Multiple finishes (Anodized Aluminum, White Powder Coated)
- Automatic size rounding to nearest available size
- With Damper (WD) option support
- Quote management (add, remove, clear items)
- Save quotes to text files
- Professional quotation formatting

## Installation

1. **Install Python 3.8+** (if not already installed)

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Ensure `price_list_modified.xlsx` is in the same folder**

## Running the Application

### As Python Script:
```bash
python main.py
```

### Build as Executable (.exe):

**Method 1: Using the build script (recommended)**
```bash
python build_exe.py
```

**Method 2: Manual PyInstaller command**
```bash
pyinstaller --onefile --windowed --name=QuotationSystem --add-data="price_list_modified.xlsx;." main.py
```

Note: On macOS/Linux, use `:` instead of `;` in the `--add-data` argument.

The executable will be created in the `dist` folder.

## Usage

1. **Launch the application**
2. **Enter customer information** (name, quote number)
3. **Select product configuration:**
   - Product type (HRG or WSG)
   - Finish
   - Width and height (in inches)
   - With Damper option
   - Quantity
4. **Click "Add to Quote"** to add items
5. **Save or print** your quote when ready

## Notes

- The system automatically rounds dimensions to the nearest available size
- Prices are loaded from the Excel file at startup
- The Excel file must be in the same directory as the application
- When packaged as .exe, the Excel file is bundled automatically

## Troubleshooting

**Error: Price list file not found**
- Ensure `price_list_modified.xlsx` is in the same directory as the application

**Error: Failed to load price list**
- Check that the Excel file is not corrupted
- Verify the sheet names match: "1-HRG,2-WSG Alu" and "1-HRG,2-WSG Wh"

**Module not found errors**
- Run: `pip install -r requirements.txt`

## Development

The code is modularized for easy maintenance:

- **price_loader.py**: All Excel parsing and price lookup logic
- **quotation_ui.py**: All PyQt5 GUI components and user interactions
- **main.py**: Simple entry point that launches the application

To modify the UI, edit `quotation_ui.py`.
To change price loading logic, edit `price_loader.py`.

