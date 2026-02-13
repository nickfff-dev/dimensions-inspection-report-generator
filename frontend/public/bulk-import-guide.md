# Bulk Item Import Guide

## Overview
The bulk import feature allows you to quickly add multiple inspection items to your release form using either CSV file upload or copy-paste from Excel/Google Sheets.

## CSV Upload Method

### Supported Format
```csv
Drawing No,Rev,Item No,Qty (Nos.),Remarks
FS131-704365-324401-324-1-B-130,05,B-130,1,Steel structure frame
FS131-704365-324401-324-1-B-131,02,B-131,2,Support brackets
```

### Required Fields
- **Drawing No**: Technical drawing reference number
- **Rev**: Revision number 
- **Item No**: Unique item identifier
- **Qty (Nos.)**: Quantity (must be positive integer)

### Optional Fields
- **Remarks**: Additional notes or comments

### Column Name Flexibility
The system accepts multiple column name variations:
- Drawing No: `Drawing No`, `drawingNo`, `drawing_no`, `drawing`
- Rev: `Rev`, `rev`, `revision`, `Rev No`
- Item No: `Item No`, `itemNo`, `item_no`, `item`
- Qty: `Qty`, `qty`, `quantity`, `Qty (Nos.)`
- Remarks: `Remarks`, `remarks`, `notes`, `comment`

## Excel Copy-Paste Method

### How to Use
1. Copy data from Excel/Google Sheets (Ctrl+C)
2. Click "Paste from Excel" button  
3. Paste data into the text area (Ctrl+V)
4. Click "Process Pasted Data"

### Expected Format
Tab-delimited data with headers in first row:
```
Drawing No	Rev	Item No	Qty (Nos.)	Remarks
FS131-704365-324401	05	B-130	1	Steel frame
FS131-704365-324401	02	B-131	2	Support brackets
```

## Import Modes

### Add Mode (Default)
- Adds imported items to existing items in the form
- Preserves current data
- Auto-assigns new serial numbers

### Replace Mode  
- Replaces all existing items with imported ones
- Clears current data completely
- Resets serial numbers starting from 1

## Validation & Error Handling

### Automatic Validation
- Required field checking
- Quantity validation (must be positive number)
- Duplicate detection
- Data type validation

### Error Preview
- Shows specific errors with row and field information
- Prevents import until all errors are fixed
- Provides clear error descriptions

## Import Preview
- Shows exactly what will be imported
- Displays serial number assignments
- Allows review before final import
- Shows item count and validation status

## Best Practices

### For CSV Files
1. Include headers in first row
2. Use consistent naming conventions
3. Ensure all required fields are populated
4. Validate quantities are positive numbers
5. Save as CSV (UTF-8) format

### For Excel Copy-Paste
1. Select headers + data rows
2. Copy entire selection (including headers)
3. Ensure no merged cells in copied area
4. Use standard Excel formatting

## Error Resolution

### Common Issues
- **Missing Required Fields**: Ensure Drawing No, Rev, Item No are provided
- **Invalid Quantities**: Must be positive integers (1, 2, 3...)  
- **Empty Rows**: System automatically skips blank rows
- **Special Characters**: Some characters may cause parsing issues

### Troubleshooting
1. Check CSV encoding (use UTF-8)
2. Verify column headers match expected names
3. Remove any special formatting from Excel
4. Ensure no formula cells in copied data
5. Check for trailing spaces in data

## Performance Notes
- Supports importing 100+ items efficiently
- Preview loads instantly for validation
- Large imports (500+ items) may take few seconds
- No limit on file size for CSV uploads

## Sample Files
- Download `sample-items.csv` from the application to see correct format
- Use as template for your own data imports