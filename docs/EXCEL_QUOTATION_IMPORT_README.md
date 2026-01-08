# Excel Quotation Import Template Documentation

This document describes the required format for Excel files used to import multiple items into QuoteFlow quotations.

## Overview

QuoteFlow's Excel import feature allows you to bulk-add items to a quotation by reading data from an Excel file. The importer is flexible and automatically detects the column layout by searching for header keywords.

## Quick Start

1. Create Excel file with required columns (at minimum: Model)
2. Add data rows with product information
3. In QuoteFlow, click "Import from Excel" button
4. Select your file and review imported items

## File Format Requirements

### Supported File Types

- **Excel 2007+**: `.xlsx` format only
- **Single Sheet**: Only the active (first) sheet is processed
- **Column Detection**: Automatic based on header keywords

### Required Columns

Only one column is absolutely required:

| Column | Required | Keywords | Description |
|--------|----------|----------|-------------|
| **Model** | Yes | "model" | Product model identifier |

### Optional Columns

These columns enhance the import but are not required:

| Column | Keywords | Description | Default if Missing |
|--------|----------|-------------|-------------------|
| **Detail** | "detail" | Additional item description | Empty string |
| **Width** | "width" | Width dimension | Must provide via UI |
| **Height** | "height" | Height dimension | Must provide via UI |
| **Unit** | "unit" | Dimension unit (inches/mm/cm/m/ft) | "inches" |
| **Quantity** | "quantity" | Number of units | 1 |
| **Finish** | "finish" | Finish type | None (multiplier = 1.0) |
| **Discount** | "discount" | Discount percentage | 0% |

### Column Detection

- **Case Insensitive**: "Model", "MODEL", "model" all work
- **Flexible Naming**: Partial matches work (e.g., "Product Model")
- **Search Range**: First 20 rows searched for header keywords
- **First Match**: Uses first occurrence if multiple matches found

## Row Types

### 1. Product Items

Regular data rows with product information:

```
| Model   | Detail              | Width | Height | Unit   | Quantity | Finish        | Discount |
|---------|---------------------|-------|--------|--------|----------|---------------|----------|
| AA-100  | Office Window       | 24    | 36     | inches | 2        | Anodized      | 0        |
| AA-200  | Conference Room     | 30    | 48     | inches | 4        | Powder Coated | 5        |
| BB-100  | Entry Door          | 36    | 84     | inches | 1        | No Finish     | 0        |
```

### 2. Title Rows (Section Headers)

Rows with Model text but no other data become section titles:

```
| Model            | Detail | Width | Height | Quantity | Finish | Discount |
|------------------|--------|-------|--------|----------|--------|----------|
| First Floor      |        |       |        |          |        |          |
| AA-100           | Office | 24    | 36     | 2        | Anod.  | 0        |
| AA-200           | Conf.  | 30    | 48     | 4        | Powder | 5        |
| Second Floor     |        |       |        |          |        |          |
| BB-100           | Entry  | 36    | 84     | 1        | None   | 0        |
```

**Title Row Detection**:
- Model column has text (not blank)
- All other important columns are blank (Detail, Width, Height, Quantity, Finish)
- Creates visual separator in quotation
- No price calculation for title rows

### 3. Blank Rows

Completely blank rows are also converted to title rows with empty text.

## Column Format Details

### Model Column

**Format**: Text string
**Examples**: 
- `AA-100`
- `Model-AA-100`
- `PRODUCT-X`

**Requirements**:
- Must match a model in the price database
- Case-sensitive matching
- No automatic correction or suggestions

**Invalid Items**:
- Non-existent models are logged as errors
- Item is skipped but import continues

### Width and Height Columns

**Format**: Numeric values
**Units**: Specified in "Unit" column
**Examples**:
```
| Width | Height | Unit       |
|-------|--------|------------|
| 24    | 36     | inches     |
| 610   | 914    | millimeters|
| 24    | 36     | mm         |
```

**Parsing**:
- Accepts numeric values with or without unit symbols
- Supports decimals: `24.5`, `36.75`
- Automatically converts mm to inches internally

**Unit Column Values**:
- `"inches"`, `"inch"`, `"in"`, `"""` (default)
- `"millimeters"`, `"millimeter"`, `"mm"`
- `"centimeters"`, `"centimeter"`, `"cm"`
- `"meters"`, `"meter"`, `"m"`
- `"feet"`, `"foot"`, `"ft"`
- Case-insensitive matching

**Missing Dimensions**:
- If Width or Height missing, uses database default or prompts user
- For some products, dimensions may be optional

### Quantity Column

**Format**: Numeric integer or decimal
**Examples**:
```
| Quantity |
|----------|
| 1        |
| 5        |
| 10       |
| 2.5      |  ← Decimals supported but unusual
```

**Parsing**:
- Converts to integer if possible
- Supports decimal quantities for special cases
- Invalid values default to 1

### Finish Column

**Format**: Text string matching finish types
**Accepted Values** (case-insensitive):
- **Anodized Aluminum**: `"Anodized"`, `"Anodized Aluminum"`, `"Aluminum"`, `"Anodised"`, `"Aluminium"`
- **Powder Coated**: `"Powder Coated"`, `"Powder"`, `"Coated"`, `"PC"`, or Thai color names (e.g., `"ขาวนวล"`, `"ดำด้าน"`)
- **No Finish**: `"No Finish"`, `"None"`, `"Raw"`, `"Unfinished"`, `""` (blank), `"สังกะสี"`, `"Stainless Steel"`
- **Special Color**: Any other finish value will be treated as a special color with custom multiplier

**Examples**:
```
| Finish          | Interpreted As         |
|-----------------|------------------------|
| Anodized        | Anodized Aluminum      |
| Anodized Aluminum | Anodized Aluminum   |
| Powder Coated   | Powder Coated           |
| No Finish       | No Finish               |
| (blank)         | No Finish               |
| Raw             | No Finish               |
| ANODIZED        | Anodized Aluminum       |
| Special Color - Red | Special Color (with multiplier) |
```

**Behavior**:
- Blank cells default to No Finish (multiplier = 1.0)
- Invalid finish types may cause import warnings
- Finish multipliers are applied from database

### Discount Column

**Format**: Numeric percentage value
**Examples**:
```
| Discount |
|----------|
| 0        |  ← No discount
| 5        |  ← 5% discount
| 10       |  ← 10% discount
| 15.5     |  ← 15.5% discount
```

**Parsing**:
- Enter as plain number (not percentage symbol)
- Value represents percentage points
- `5` means 5% off, not 0.05% off
- Negative values not supported

**Calculation**:
- Applied after base price and finish multiplier
- Discount = (Price × Discount / 100)
- Discounted Price = Price - Discount

## Complete Example

### Minimal Format (Model Only)

```
| Model   |
|---------|
| AA-100  |
| AA-200  |
| BB-100  |
```

Result: Items imported with default dimensions, quantities, and finishes.

### Standard Format (All Columns)

```
| Model  | Detail           | Width | Height | Unit   | Quantity | Finish        | Discount |
|--------|------------------|-------|--------|--------|----------|---------------|----------|
| AA-100 | Main Office      | 24    | 36     | inches | 2        | Anodized      | 0        |
| AA-100 | Conference Room  | 30    | 48     | inches | 3        | Powder Coated | 5        |
| AA-200 | Break Room       | 20    | 30     | inches | 1        | Anodized      | 0        |
| BB-100 | Entry Vestibule  | 36    | 84     | inches | 2        | No Finish     | 10       |
```

### With Titles and Sections

```
| Model          | Detail           | Width | Height | Unit   | Qty | Finish        | Disc |
|----------------|------------------|-------|--------|--------|-----|---------------|------|
| FIRST FLOOR    |                  |       |        |        |     |               |      |
| Office Area    |                  |       |        |        |     |               |      |
| AA-100         | Office Window    | 24    | 36     | inches | 2   | Anodized      | 0    |
| AA-100         | Office Door      | 36    | 84     | inches | 1   | Powder Coated | 5    |
|                |                  |       |        |        |     |               |      |
| SECOND FLOOR   |                  |       |        |        |     |               |      |
| Conference     |                  |       |        |        |     |               |      |
| AA-200         | Main Conf Room   | 48    | 60     | inches | 3   | Anodized      | 0    |
| BB-100         | Small Conf Room  | 30    | 48     | inches | 2   | No Finish     | 10   |
```

## Import Process

### Step-by-Step Flow

1. **File Selection**: User clicks "Import from Excel" and selects file
2. **Header Detection**: Searches first 20 rows for column keywords
3. **Data Reading**: Processes all rows below header
4. **Item Parsing**: Converts each row to item object
5. **Validation**: Checks model exists, dimensions valid, etc.
6. **Price Calculation**: Calculates prices for valid items
7. **Progress Display**: Shows progress bar during import
8. **Results Summary**: Displays success count, errors, warnings

### Progress Reporting

During import, you'll see:
```
Loading Excel file...                    [5%]
Searching for header row...              [10%]
Reading row 15 of 50...                  [35%]
Processing items...                      [70%]
Calculating prices...                    [85%]
Parsing complete!                        [100%]
```

### Import Results Dialog

After import completes:

```
Upload Complete!

Successfully imported 12 item(s) and 3 title(s)

2 item(s) with errors and 1 warning(s)
```

**Errors Section**:
```
=== ERRORS ===

1. Model: INVALID-MODEL
   Error: Product 'INVALID-MODEL' not found in database

2. Model: AA-100
   Error: Invalid dimensions: width or height missing
```

**Warnings Section**:
```
=== WARNINGS ===

1. Row 15: Quantity missing, using default value of 1
```

## Validation and Error Handling

### Valid Items

Items are successfully imported when:
- Model exists in price database
- Dimensions are valid numbers (if provided)
- Quantity is valid number (or defaults to 1)
- Finish is recognized (or defaults to No Finish)
- Unit is recognized (or defaults to inches)

### Invalid Items

Items are skipped with errors when:
- Model not found in database
- Required dimensions missing (product requires them)
- Invalid numeric values for dimensions
- Critical parsing errors

**Behavior**:
- Invalid items are logged but don't stop import
- Other valid items still get imported
- Error details shown in results dialog
- User can review and manually add invalid items

### Warnings

Warnings are generated for:
- Missing optional fields (using defaults)
- Unusual dimension values (very large/small)
- Unrecognized finish types (using default)
- Quantity = 0 (unusual but allowed)

## Best Practices

### File Organization

1. **Use Clear Headers**: Make column names obvious
2. **Consistent Units**: Use same unit throughout (preferably inches)
3. **Group by Area**: Use title rows to organize by location
4. **Add Details**: Use Detail column for room names, notes
5. **Verify Models**: Double-check model names match database

### Data Entry

1. **Complete Information**: Fill all columns when possible
2. **Consistent Formatting**: Same finish names throughout
3. **Numeric Precision**: Use appropriate decimal places
4. **Test Small First**: Try importing 5-10 items before full list
5. **Save Backup**: Keep copy before making major changes

### Error Prevention

1. **Check Model List**: Verify models exist in QuoteFlow first
2. **Use Templates**: Start from working example
3. **Avoid Formulas**: Use values only (not Excel formulas)
4. **No Merged Cells**: Keep simple table structure
5. **Clean Data**: Remove extra spaces, special characters

## Troubleshooting

### Problem: Header Not Found

**Symptoms**: "Could not find required 'Model' column"
**Causes**:
- Column header doesn't contain word "model"
- Header row beyond first 20 rows
- Column header is in merged cell

**Solutions**:
- Rename column to include "Model" keyword
- Move header to top 20 rows
- Unmerge header cells

### Problem: All Items Show as Invalid

**Symptoms**: Every item reports "Product not found"
**Causes**:
- Model names don't match database
- Database not generated correctly
- Wrong product sheet selected

**Solutions**:
- Check spelling of model names
- Regenerate database from Excel price list
- Verify models exist in Header sheet
- Match case exactly (case-sensitive)

### Problem: Dimensions Not Working

**Symptoms**: "Invalid dimensions" errors
**Causes**:
- Non-numeric values in Width/Height columns
- Cells contain formulas showing errors
- Unit column not recognized

**Solutions**:
- Use "Paste Values" to remove formulas
- Check for text in numeric columns
- Use recognized unit keywords
- Try leaving Unit column blank (defaults to inches)

### Problem: Prices Incorrect

**Symptoms**: Imported items have wrong prices
**Causes**:
- Wrong finish selected
- Discount not applied correctly
- Database has wrong prices
- Base modifier equation incorrect

**Solutions**:
- Verify finish names match exactly
- Check discount as percentage (not decimal)
- Regenerate database from latest Excel
- Review finish multipliers in database Header sheet

### Problem: Slow Import

**Symptoms**: Import takes several minutes
**Causes**:
- Very large file (>1000 rows)
- Complex calculations
- Many invalid items requiring validation

**Solutions**:
- Split into smaller files
- Remove unnecessary columns
- Pre-validate model names
- Close other applications

### Problem: Titles Not Working

**Symptoms**: Title rows become items or errors
**Causes**:
- Model text matches actual product
- Other columns have data
- Formatting issue

**Solutions**:
- Use clearly non-product text for titles
- Ensure Detail/Width/Height/Quantity/Finish are blank
- Use ALL CAPS or special characters for titles

## Advanced Features

### Multiple Imports

You can import from multiple Excel files:
1. Import first file
2. Review items in quotation
3. Import second file (adds to existing items)
4. Continue as needed

Items from different imports are combined in the same quotation.

### Import and Edit

After importing:
1. Items appear in quotation table
2. Edit any item (dimensions, quantity, finish, discount)
3. Delete unwanted items
4. Add manual items between imported ones
5. Export final quotation

### Template Files

Create reusable templates:
1. Set up Excel with common columns and formatting
2. Save as template file
3. Copy template for each new project
4. Fill in specific project data
5. Import into QuoteFlow

### Data Validation in Excel

For better reliability, use Excel data validation:
- **Finish Column**: Dropdown list (Anodized Aluminum, Powder Coated, No Finish, Special Color)
- **Unit Column**: Dropdown list (inches, millimeters, centimeters, meters, feet)
- **Quantity**: Number validation (>0)
- **Discount**: Number validation (0-100)

## Integration with Other Systems

### Exporting from CAD

If exporting from CAD software:
1. Generate schedule/list with dimensions
2. Add Model column with appropriate values
3. Include quantity counts
4. Save as Excel format
5. Import into QuoteFlow

### Spreadsheet Tools

Compatible with:
- Microsoft Excel (2007+)
- Google Sheets (export as .xlsx)
- LibreOffice Calc (save as .xlsx)
- Numbers (export as .xlsx)

**Note**: Always save/export as `.xlsx` format before importing.

## Performance Tips

- **Batch Processing**: Import 100-200 items at a time for best performance
- **Clean Data**: Remove empty rows and columns
- **Simple Format**: Avoid complex formatting, merged cells, formulas
- **Close File**: Close Excel file before importing
- **Save Changes**: Save Excel file before selecting for import

## Sample Files

QuoteFlow includes sample files in the project:
- Check `src/` folder for `Quotation_25-12023.xlsx` as an example export format
- Use as reference for structure and formatting

---

**Related Documentation**:
- Main README: Project overview and setup
- [Excel Price Template](EXCEL_PRICE_TEMPLATE_README.md): Price list format for database generation

**Support**: For issues with Excel imports, verify your format matches this documentation and check the results dialog for specific error messages.
