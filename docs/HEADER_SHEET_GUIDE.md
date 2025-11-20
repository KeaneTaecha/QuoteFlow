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
| **Base Modifier** (formerly **TB Modifier**) | Equation or multiplier to calculate base price (BP) from table price | "TB*1.2", "TB+50", "1.5" |
| **Anodized Multiplier** | Price multiplier for Anodized Aluminum finish | 1.2, 1.3, etc. |
| **Powder Coated Multiplier** | Price multiplier for White Powder Coated finish | 1.35, 1.4, etc. |
| **WD** | With Damper equation or multiplier (can reference BP) | "BP*1.1", "TB+BP*0.5", "1.2" |

### Example Header Sheet

```
| Table id | Sheet           | Model          | Base Modifier | Anodized Multiplier | Powder Coated Multiplier | WD |
|----------|-----------------|----------------|-------------|---------------------|-------------------------|-----|
| 1        | 1-HRG,2-WSG TB  | 1-HRG, 2-WSG  | TB*1.1      | 1.2                 | 1.35                    | BP*1.05 |
| 2        | 1-AL, 1-RAL Alu | 1-AL          | 1.15        | 1.25                | 1.4                     | 1.1 |
| 3        | 1-PFR TB        | 1-PFR         | TB+25       | 1.3                 | 1.45                    | BP*1.08 |
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

## Price Calculation Flow

### 1. **Base Modifier (Base Price Calculation)**
The Base Modifier (formerly TB Modifier) is applied first to calculate the base price (BP) from the table price (TB):
- **Base Modifier**: Can be a simple multiplier (e.g., "1.2") or an equation (e.g., "TB*1.1", "TB+50")
- **Result**: Creates the Base Price (BP) which can be referenced in other equations
- **Variables Available**: TB (table price), WD (with damper price), WIDTH, HEIGHT

### 2. **WD (With Damper) Calculation**
The WD column can reference the calculated BP and creates MWD (Modified WD):
- **WD Equation**: Can use BP, TB, WD, WIDTH, HEIGHT variables
- **Examples**: "BP*1.05", "TB+BP*0.5", "1.1"
- **Result**: Creates MWD (Modified WD) price used for with-damper calculations

### 3. **Finish Multipliers**
The system applies finish multipliers to the calculated prices:
- **Anodized Aluminum**: Uses the Anodized Multiplier from Header sheet
- **White Powder Coated**: Uses the Powder Coated Multiplier from Header sheet  
- **Other Paint**: Uses the Other Paint Multiplier from Header sheet

These multipliers are applied when loading prices from the database. The multipliers are configured per product in the Header sheet.

## Available Variables in Equations

When writing equations for Base Modifier or WD columns, you can use these variables:

| Variable | Description | Example Value |
|----------|-------------|---------------|
| **TB** | Original table price (unchanged) | 100 |
| **WD** | Original with-damper price (unchanged) | 120 |
| **BP** | Base price calculated from Base Modifier | 110 (if Base Modifier = "TB*1.1") |
| **MWD** | Modified WD price calculated from WD equation | 130 (if WD = "BP*1.18") |
| **WIDTH** | Product width in inches | 24 |
| **HEIGHT** | Product height in inches | 36 |

### Example Calculation Flow:
1. TB = 100 (original table price)
2. Base Modifier = "TB*1.1" → BP = 110
3. WD Equation = "BP*1.18" → MWD = 129.8
4. Finish Multiplier = 1.5 → Final Price = 194.7 (when with_damper=True)

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

