# Header Sheet Guide

## Overview
The quotation system now uses a **Header sheet** in the Excel file to dynamically configure products and their price tables. This allows flexible management of product models without code changes.

## Header Sheet Structure

The Header sheet must contain the following columns:

| Column | Description | Example |
|--------|-------------|---------|
| **Table id** | Unique identifier for each price table | 1, 2, 3, etc. |
| **Sheet** | Name of the Excel sheet containing the price table | "1-HRG,2-WSG TB" |
| **Model** | Product model(s) for this table (comma-separated) | "1-HRG, 2-WSG" or "1-HRG" |

### Example Header Sheet

```
| Table id | Sheet           | Model          |
|----------|-----------------|----------------|
| 1        | 1-HRG,2-WSG TB  | 1-HRG, 2-WSG  |
| 2        | 1-AL, 1-RAL Alu | 1-AL          |
| 3        | 1-PFR TB        | 1-PFR         |
```

## How It Works

### 1. **Single Model Configuration**
When a table has only ONE model:
```
Table id: 1
Sheet: 1-HRG TB
Model: 1-HRG
```
- The UI will show **only "1-HRG"** in the product dropdown
- Prices will be loaded from the "1-HRG TB" sheet
- Product code will be "1-HRG"

### 2. **Multiple Models Configuration**
When a table has MULTIPLE models (separated by commas):
```
Table id: 1
Sheet: 1-HRG,2-WSG TB
Model: 1-HRG, 2-WSG
```
- The UI will show **both "1-HRG" and "2-WSG"** in the product dropdown
- Both models share the same price table from "1-HRG,2-WSG TB" sheet
- Each model can be selected independently in the UI

### 3. **Database Structure**

The system creates two tables in SQLite:

#### Products Table
```sql
CREATE TABLE products (
    product_id INTEGER PRIMARY KEY AUTOINCREMENT,
    table_id INTEGER NOT NULL,
    model TEXT NOT NULL,
    sheet_name TEXT NOT NULL,
    UNIQUE(table_id, model)
)
```

Example data:
| product_id | table_id | model  | sheet_name      |
|------------|----------|--------|-----------------|
| 1          | 1        | 1-HRG  | 1-HRG,2-WSG TB |
| 2          | 1        | 2-WSG  | 1-HRG,2-WSG TB |

#### Prices Table
```sql
CREATE TABLE prices (
    price_id INTEGER PRIMARY KEY AUTOINCREMENT,
    table_id INTEGER NOT NULL,
    width INTEGER NOT NULL,
    height INTEGER NOT NULL,
    normal_price REAL,
    price_with_damper REAL,
    UNIQUE(table_id, width, height)
)
```

## Workflow

### Step 1: Update Header Sheet in Excel
1. Open `price_list_modified.xlsx`
2. Go to the **Header** sheet
3. Add a new row with:
   - Table id (unique number)
   - Sheet name (exact name of the price table sheet)
   - Model(s) (single model or comma-separated list)

### Step 2: Run Conversion Script
```bash
python getsql.py
```

This will:
- Read the Header sheet
- Extract table metadata (id, sheet name, models)
- Parse models (split by comma if multiple)
- Create product records for each model
- Load prices from the specified sheet
- Link prices to products via table_id

### Step 3: Run the Application
```bash
python main.py
```

The UI will:
- Load all models from the database
- Populate the product dropdown dynamically
- Show only the models defined in the Header sheet

## Benefits

1. **Flexible Configuration**: Add/remove products by editing the Header sheet
2. **No Code Changes**: Product models are loaded dynamically from the database
3. **Shared Price Tables**: Multiple models can share the same price data
4. **Easy Maintenance**: All product configuration in one central location

## Example Scenarios

### Scenario 1: Add a New Single-Model Product
```
Table id: 2
Sheet: 1-PFR TB
Model: 1-PFR
```
Result: Only "1-PFR" appears in the UI dropdown

### Scenario 2: Add a Product with 3 Models
```
Table id: 3
Sheet: Diffusers
Model: Model-A, Model-B, Model-C
```
Result: All three models appear in the UI dropdown, sharing the same prices

### Scenario 3: Current Configuration
```
Table id: 1
Sheet: 1-HRG,2-WSG TB
Model: 1-HRG, 2-WSG
```
Result: Both "1-HRG" and "2-WSG" appear in the UI dropdown

## Price Multipliers

The system applies finish multipliers to base prices:
- **Anodized Aluminum**: 1.2x base price
- **White Powder Coated**: 1.35x base price

These multipliers are applied when loading prices from the database.

## Troubleshooting

### Issue: No products shown in UI
- Check that the Header sheet exists in the Excel file
- Verify Header sheet has data (not just headers)
- Run `python getsql.py` to regenerate the database

### Issue: Wrong products showing
- Check the Model column in Header sheet
- Ensure model names are spelled correctly
- Verify comma separation for multiple models

### Issue: No prices for a product
- Verify the Sheet name in Header matches the actual sheet name
- Check that the price table sheet has data
- Ensure table_id is unique for each entry

