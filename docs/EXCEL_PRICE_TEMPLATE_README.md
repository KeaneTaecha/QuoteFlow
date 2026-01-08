# Excel Price List Template Documentation

This document describes the required format for Excel price list files that are imported into QuoteFlow's SQLite database using `getsql.py`.

## Overview

The Excel price list contains product pricing information organized into multiple sheets and tables. QuoteFlow's converter automatically detects and processes three types of table formats:

1. **Header Sheet**: Product metadata, models, and multipliers
2. **Standard Tables**: Traditional price matrices with dimensions
3. **Other Tables**: Alternative price matrices with flexible column layouts

## Quick Start

1. **Required Sheet**: Must have a sheet named "Header" (case-insensitive)
2. **Data Sheets**: Additional sheets containing price tables
3. **Run Converter**: `python src/excel_to_sql/getsql.py` to generate `prices.db`

## Sheet Types

### 1. Header Sheet (Required)

The Header sheet contains product metadata and finish multipliers. This sheet must exist and should be named "Header" (case-insensitive).

#### Required Columns

The converter uses keyword recognition to identify columns. Column names are case-insensitive and can include these keywords:

| Column Purpose | Recognized Keywords |
|---------------|-------------------|
| **Sheet Name** | "sheet", "sheet name", "sheet_name", "sheetname" |
| **Model** | "model", "models", "product", "product model" |
| **Base Modifier** | "tb modifier", "tb_modifier", "tb modifier equation", "tb_modifier_equation", "base modifier", "base_modifier", "base price modifier", "base_price_modifier", "bp modifier", "bp_modifier" |
| **Anodized Multiplier** | "anodized", "aluminum", "anodized aluminum", "anodized multiplier", "anodized aluminum multiplier" |
| **Powder Coated Multiplier** | "powder coated", "powder_coated", "powder", "coated", "powder coated multiplier", "powdercoated" |
| **No Finish Multiplier** | "no finish", "no_finish", "no finish multiplier", "no_finish_multiplier", "raw", "raw multiplier", "unfinished", "unfinished multiplier" |
| **WD Multiplier** | "wd", "with damper", "with_damper", "wd multiplier", "wd equation", "damper", "damper multiplier", "damper equation" |

#### Example Header Sheet Format

```
| Sheet Name | Model    | TB Modifier | Anodized | Powder Coated | No Finish | WD      |
|-----------|---------|-------------|----------|---------------|-----------|---------|
| Sheet1    | AA-100  | 1.0         | 1.0      | 1.15          | 0.85      | 1.2     |
| Sheet1    | AA-200  | w*h/144     | 1.0      | 1.20          | 0.80      |         |
| Sheet2    | BB-100  | (w+h)/2     | 1.05     | 1.15          | 0.85      | 1.25    |
```

#### Column Details

**Sheet Name**:
- Specifies which sheet contains prices for this model
- Must match an existing sheet name in the workbook
- Case-sensitive matching

**Model**:
- Unique product model identifier
- Used throughout the application
- Should match models in quotations

**TB Modifier (Base Modifier)**:
- Mathematical equation or numeric value
- Applied before finish multipliers
- Supports: `w` (width), `h` (height), `+`, `-`, `*`, `/`, `(`, `)`
- Examples:
  - `1.0` - Fixed multiplier
  - `w*h/144` - Area in square feet
  - `(w+h)/2` - Average dimension
  - `w*1.5` - 50% width premium

**Finish Multipliers**:
- **Anodized**: Typically 1.0 (standard finish)
- **Powder Coated**: Usually > 1.0 (premium finish)
- **No Finish**: Usually < 1.0 (raw material)
- Leave blank for default value of 1.0

**WD Multiplier (With Damper)**:
- Additional cost multiplier when damper is selected
- Leave blank if damper not applicable
- Typically > 1.0 to account for damper cost

### 2. Standard Table Format

Standard tables are the most common format with width columns and height rows.

#### Structure

```
[Model Name]
       4"      6"      8"      10"     12"
4"    $10.50  $12.00  $14.00  $16.00  $18.00
6"    $12.00  $14.00  $16.50  $19.00  $22.00
8"    $14.00  $16.50  $19.50  $23.00  $27.00
10"   $16.00  $19.00  $23.00  $28.00  $33.00
12"   $18.00  $22.00  $27.00  $33.00  $40.00
```

#### Format Requirements

1. **Model Name Row**: Text identifying the product (optional, can be on row above table)
2. **Width Header Row**: Contains width values with inch symbol (`"`)
3. **Height Rows**: Start with height value followed by prices
4. **Price Format**: Numeric values (dollar signs optional)
5. **Spacing**: Adjacent rows (no blank rows between data rows) OR alternating rows (blank row between each data row)

#### Detection Rules

- Searches for inch values (`4"`, `6"`, etc.) in the first few rows
- Width row contains horizontal inch values
- Height column contains vertical inch values
- Automatically detects separated or adjacent row format

#### Example Variations

**Adjacent Rows (Compact)**:
```
      4"    6"    8"
4"   $100  $120  $140
6"   $120  $145  $170
8"   $140  $170  $200
```

**Separated Rows (Spaced)**:
```
      4"    6"    8"
4"   $100  $120  $140

6"   $120  $145  $170

8"   $140  $170  $200
```

### 3. Other Table Format

Other tables support alternative structures including non-inch widths or flexible column layouts.

#### Structure

Similar to Standard tables but with more flexibility:

```
[Model Name]
         Standard  Premium  Deluxe
Small    $100     $120     $140
Medium   $150     $180     $210
Large    $200     $240     $280
```

#### Format Requirements

1. **First Column**: Can contain inch values OR text labels
2. **Header Row**: Column names (can be model names, sizes, or other identifiers)
3. **Data Rows**: Prices corresponding to row/column intersections
4. **Mixed Types**: Supports combination of inch values and text labels

#### Use Cases

- Price variations by model name columns
- Size categories (Small/Medium/Large)
- Special pricing structures
- Non-dimensional pricing

#### Example: Model-Based Pricing

```
          AA-100  AA-200  AA-300
6"       $100    $110    $120
8"       $130    $145    $160
10"      $165    $185    $205
```

## Dimension Format

All dimension values should include the inch symbol (`"`):

- **Correct**: `4"`, `6"`, `8"`, `10"`, `12"`
- **Incorrect**: `4`, `6`, `8` (without inch symbol)
- **Decimal Support**: `7.2"`, `8.5"` are supported
- **Merged Cells**: Supported for repeated dimensions

## Price Format

Prices can be entered in multiple formats:

- **With Dollar Sign**: `$100.50`, `$1,250.00`
- **Without Dollar Sign**: `100.50`, `1250.00`
- **With Commas**: `$1,250.00`, `1,250.00`
- **Simple Numbers**: `100`, `250.5`

All formats are automatically parsed to numeric values.

## Multi-Table Detection

QuoteFlow automatically detects multiple tables on a single sheet:

1. **Table Identification**: Looks for model names or dimension patterns
2. **Boundary Detection**: Determines where each table starts and ends
3. **Separation**: Processes each table independently
4. **Unique IDs**: Assigns unique `table_id` to each detected table

### Multiple Tables Per Sheet Example

```
Sheet: "Products"

Model: AA-100
      4"    6"    8"
4"   $100  $120  $140
6"   $120  $145  $170
[blank rows]

Model: AA-200
      4"    6"    8"
4"   $110  $135  $155
6"   $135  $165  $190
[blank rows]

Model: BB-100
         Standard  Premium
Small    $50      $65
Medium   $75      $95
```

All three tables would be detected and processed automatically.

## Row Multipliers (Advanced)

For products that require width-specific multipliers:

### Structure

Row multipliers can be included in the same table or separately:

```
[Model with Row Multipliers]
        4"    6"    8"
        1.0   1.05  1.10   ← Row multipliers
4"     $100  $110  $125
6"     $120  $135  $155
```

### Purpose

- Apply additional multiplier based on width
- Separate multipliers for regular and WD (with damper) configurations
- Stored in `row_multipliers` table in database

## Filter Pricing

Special tables for filter products:

### Format

```
[Filter Model]
2-Inch     $15.00
4-Inch     $25.00
Activated  $35.00
Poly       $20.00
```

### Supported Filter Types

- `2"` or `2-inch` or `2 inch`
- `4"` or `4-inch` or `4 inch`
- `Activated Carbon` or `Activated` or `AC`
- `Poly` or `Polyester`

## Database Generation

### Running the Converter

```bash
cd src/excel_to_sql
python getsql.py
```

Or from project root:

```bash
python src/excel_to_sql/getsql.py
```

### Converter Features

1. **Automatic Detection**: Finds all sheets and tables automatically
2. **Smart Headers**: Uses keyword recognition for column identification
3. **Multi-Table Support**: Processes multiple tables per sheet
4. **Error Reporting**: Shows warnings for skipped data
5. **Statistics**: Displays counts of processed items

### Output

Creates `prices.db` SQLite database with tables:
- `products`: Product metadata and multipliers
- `prices`: Price matrices
- `row_multipliers`: Width-specific multipliers for height exceeded scenarios
- `column_multipliers`: Height-specific multipliers for width exceeded scenarios

## Validation and Error Checking

The converter performs automatic validation:

### Common Warnings

1. **Missing Model in Header**: Product found in data sheets but not in Header sheet
2. **Invalid Dimensions**: Non-numeric or malformed dimension values
3. **Duplicate Entries**: Same product/dimension combination appears twice
4. **Missing Prices**: Gaps in price matrix
5. **Sheet Not Found**: Header references sheet that doesn't exist

### Best Practices

1. Always maintain a Header sheet with all products
2. Use consistent inch symbol (`"`) for dimensions
3. Ensure no duplicate model names
4. Fill all cells in price matrices (no gaps)
5. Use clear model naming conventions
6. Test with `getsql.py` after making changes
7. Keep backup of Excel file before major changes

## Example Complete Price List

Here's a minimal complete example:

**Sheet: "Header"**
```
| Sheet Name | Model   | TB Modifier | Anodized | Powder Coated | No Finish |
|-----------|---------|-------------|----------|---------------|-----------|
| Products  | AA-100  | 1.0         | 1.0      | 1.15          | 0.85      |
| Products  | AA-200  | w*h/144     | 1.0      | 1.20          | 0.80      |
```

**Sheet: "Products"**
```
AA-100
      4"    6"    8"
4"   $100  $120  $140
6"   $120  $145  $170
8"   $140  $170  $200


AA-200
      4"    6"    8"
4"   $110  $135  $155
6"   $135  $165  $190
8"   $155  $190  $220
```

## Troubleshooting

### Problem: Tables Not Detected

**Symptoms**: Converter skips sheets or reports 0 tables
**Solutions**:
- Ensure dimension values have inch symbol (`"`)
- Check for blank rows between header and data
- Verify model names are clearly separated from price data
- Look for merged cells that might confuse detection

### Problem: Incorrect Prices in Database

**Symptoms**: Prices don't match Excel file
**Solutions**:
- Check for formula errors in Excel (use "Values Only")
- Verify no hidden rows/columns
- Ensure currency formatting isn't affecting values
- Check for merged cells in price areas

### Problem: Models Not Found

**Symptoms**: Quotations can't find products
**Solutions**:
- Verify model names match exactly (case-sensitive)
- Check Header sheet includes all models
- Ensure Sheet Name column matches actual sheet names
- Regenerate database after Excel changes

### Problem: Finish Multipliers Not Working

**Symptoms**: All finishes have same price
**Solutions**:
- Verify Header sheet has multiplier columns
- Check multiplier values are numeric
- Ensure column headers match recognized keywords
- Confirm model exists in Header sheet

## Advanced Features

### Custom Modifier Equations

Supported operators and variables:

**Variables**:
- `w` - Width in inches
- `h` - Height in inches

**Operators**:
- `+` - Addition
- `-` - Subtraction
- `*` - Multiplication
- `/` - Division
- `(` `)` - Grouping

**Examples**:
```
w*h/144          # Square feet (144 sq in = 1 sq ft)
(w+h)/2          # Average dimension
w*1.5+h*2        # Weighted dimensions
(w*h/100)+5      # Area with base fee
```

### Vertically Merged Cells

The converter supports vertically merged cells for repeating dimensions:

```
      4"    6"    8"
4"   $100  $120  $140
     $105  $125  $145  ← Same height (4") continued
6"   $120  $145  $170
```

### Empty Cell Handling

- Empty price cells are skipped (not stored)
- Empty dimension cells break table detection
- Empty multiplier cells default to 1.0
- Empty equations treated as 1.0

## Performance Considerations

- **Large Files**: Files with 1000+ products may take several minutes
- **Complex Formulas**: Equation parsing adds minimal overhead
- **Multiple Sheets**: No significant impact on processing time
- **Database Size**: Typically < 10 MB for standard price lists

## Version Compatibility

- **Excel Format**: `.xlsx` (Excel 2007 and later)
- **Not Supported**: `.xls` (older Excel format)
- **Read-Only Mode**: File is opened read-only (safe to run while file is open)

---

**Next Steps**: After creating your Excel price list following this format, run `python src/excel_to_sql/getsql.py` to generate the database, then refer back to the main README.md for application usage.
