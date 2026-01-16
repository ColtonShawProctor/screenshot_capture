# Table Boundary Detection Fix

## Problem Summary

The Excel table screenshot service was cutting off tables prematurely:
1. **Missing rows at bottom**: Tables were ending before totals/summary rows
2. **Missing columns on right**: Tables were missing data columns (especially dollar values)

## Root Causes

1. **Fixed Width Limitation**: The code used hardcoded fixed widths for known tables, which didn't account for tables that extend beyond those widths
2. **Early Row Termination**: The algorithm stopped at 2 consecutive empty rows, which could occur before totals rows
3. **Limited Scanning**: Only scanned 10 columns and 30 rows, which wasn't enough for larger tables
4. **Header-Only Column Detection**: Column width was determined only from the header row, missing data columns that extend beyond headers

## Solution Implemented

### 1. Dynamic Column Detection

**Before:**
- Used fixed widths from a dictionary (e.g., "Sources and Uses" = 7 columns)
- For unknown tables, only scanned 10 columns from header row

**After:**
- Scans up to **50 columns** (increased from 10)
- Scans **multiple rows** (first 50 rows) to find the widest point of the table
- Stops only after **3 consecutive empty columns** (instead of 1)
- This ensures data columns beyond the header are captured

### 2. Dynamic Row Detection

**Before:**
- Scanned only 30 rows
- Stopped at 2 consecutive empty rows
- No special handling for totals

**After:**
- Scans up to **100 rows** (increased from 30)
- Stops only after **3 consecutive empty rows** (instead of 2)
- **Total row detection**: When a row contains keywords like "total", "sum", "subtotal", it scans 2-3 more rows ahead to catch additional summary rows
- Adds a 1-row buffer at the end to ensure borders/formatting aren't cut off

### 3. Enhanced Content Detection

**New `cell_has_content()` function:**
- Checks for string content (text)
- Checks for numeric values (including zero, but handles it appropriately)
- Checks for formulas
- More robust than just checking `getString().strip()`

### 4. Removed Fixed Width Dictionary

The fixed width dictionary has been completely removed. All tables now use dynamic detection, which:
- Adapts to actual table sizes
- Handles tables that vary in width
- Works for unknown table types

## Key Changes in Code

### File: `capture_table.py`

**Removed:**
```python
fixed_widths = {
    'Sources and Uses': 7,
    'Take Out Loan Sizing': 3,
    'Capital Stack at Closing': 7,
    'Loan to Cost': 8,
    'Loan to Value': 7,
    'PILOT Schedule': 8,
}
```

**Added:**
- `cell_has_content()` helper function
- Multi-row column width detection
- Total row keyword detection and lookahead
- Increased scan limits (50 cols, 100 rows)
- Consecutive empty cell/row thresholds (3 instead of 1-2)

## Expected Improvements

### Sources and Uses Table
- **Before**: Rows 4-21, missing "Total Sources" and "Total Uses" rows
- **After**: Should include all rows through totals, dynamically detected width

### Capital Stack / Release at Closing Table
- **Before**: Only labels visible, missing dollar value columns
- **After**: Should include all data columns, detected by scanning multiple rows

### All Tables
- More accurate boundaries based on actual content
- Totals and summary rows included
- All data columns captured
- Better handling of tables with gaps or formatting

## Testing Recommendations

After deployment, verify these tables capture correctly:

1. **Sources and Uses**
   - ✅ Should include "Total Sources" and "Total Uses" rows
   - ✅ Should include all data columns

2. **Release at Closing / Capital Stack**
   - ✅ Should include all dollar value columns on the right
   - ✅ Should show both labels and values

3. **Loan to Cost**
   - ✅ Should include all percentage columns
   - ✅ Should include summary row

4. **Loan to Value**
   - ✅ Should include all data columns
   - ✅ Should include summary row

5. **Take Out Loan Sizing**
   - ✅ Should include "Debt Yield" and "Bridge to Exit Cushion" rows
   - ✅ Should include all columns

6. **PILOT Schedule**
   - ✅ Should include all 13+ years of data
   - ✅ Should include totals row

## Log Output Changes

**Before:**
```
Expanding 'Sources and Uses' from row 4, col 1
Using fixed width: 7 columns
Final range: cols 1-7, rows 4-21
```

**After:**
```
Expanding 'Sources and Uses' from row 4, col 1
Detected width: 9 columns (cols 1-9)
Detected height: 25 rows (rows 4-28)
Final range: cols 1-9, rows 4-28
```

The new logs show:
- Dynamically detected width (not fixed)
- Dynamically detected height
- More accurate ranges

## Performance Considerations

- The algorithm now scans more cells (50 columns × 100 rows = 5,000 cells max)
- However, it stops early when it finds boundaries (3 consecutive empty)
- In practice, most tables are much smaller, so performance impact should be minimal
- The increased accuracy is worth the small performance cost

## Backward Compatibility

- All existing table names still work
- The API endpoint (`/detect-and-capture`) is unchanged
- No breaking changes to request/response format
- Only the internal detection logic has changed

