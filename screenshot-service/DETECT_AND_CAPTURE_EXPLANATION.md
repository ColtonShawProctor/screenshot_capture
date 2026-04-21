# How `/detect-and-capture` Endpoint Works

## Overview

The `/detect-and-capture` endpoint automatically finds Excel tables by their header text and captures them as PNG images. This document explains the search mechanism, scanning behavior, and boundary detection logic.

## Flow Diagram

```
POST /detect-and-capture
    ↓
server.js: captureTable()
    ↓
Spawns Python process: capture_table.py
    ↓
capture_table.py: find_table_in_all_sheets()
    ↓
For each sheet:
    - Create search descriptor
    - Search for header text
    - If found → expand_to_table()
    - Return first match found
```

## 1. How Header Search Works

### Search Method
- **Location**: `capture_table.py`, lines 50-82 (`find_table_in_all_sheets()`)
- **Mechanism**: Uses LibreOffice UNO API's `createSearchDescriptor()` and `findFirst()`
- **Search Type**: Case-insensitive substring matching
- **Scope**: Searches **ALL sheets** sequentially

### Key Code (lines 63-68):
```python
# Create search descriptor
search = sheet.createSearchDescriptor()
search.SearchString = header_text
search.SearchCaseSensitive = False

found = sheet.findFirst(search)
```

### Important Limitations:

1. **First Match Only**: `findFirst()` returns only the **FIRST occurrence** of the text in each sheet
2. **Exact Substring Match**: The search looks for the exact substring (case-insensitive). If the Excel cell contains:
   - "Loan-to-Cost" (with hyphen) → won't match "Loan to Cost" (with spaces)
   - "Loan  to  Cost" (double spaces) → might not match "Loan to Cost" (single space)
   - "Loan to Cost:" (with colon) → won't match "Loan to Cost" (without colon)
3. **Stops at First Sheet**: Once a match is found in any sheet, the function returns immediately (line 76)
4. **No Fuzzy Matching**: There's no fuzzy matching or partial word matching

### What Gets Scanned:
- **Entire sheet content**: LibreOffice's search scans all cells in the sheet
- **Search order**: Typically left-to-right, top-to-bottom (LibreOffice's default behavior)
- **Merged cells**: If the header is in a merged cell, the search should still find it

## 2. Which Columns/Rows Are Scanned

### During Header Search:
- **All columns and rows** in each sheet are scanned by LibreOffice's search
- No column/row limits during the initial search phase

### After Header Found:
- The search returns the **cell address** where the header was found
- This becomes the starting point (`start_col`, `start_row`) for boundary detection

## 3. How Table Boundaries Are Determined

### Location: `expand_to_table()` function (lines 85-151)

### Width Detection:

**For Known Tables** (lines 94-101, 106-108):
- Uses a **fixed width dictionary**:
  ```python
  fixed_widths = {
      'Sources and Uses': 7,
      'Take Out Loan Sizing': 3,  # J, K, L columns
      'Capital Stack at Closing': 7,
      'Loan to Cost': 8,
      'Loan to Value': 7,
      'PILOT Schedule': 8,
  }
  ```
- If the table name matches, it uses the fixed width: `end_col = start_col + fixed_widths[table_name] - 1`

**For Unknown Tables** (lines 109-120):
- Scans up to **10 columns to the right** from the header cell
- Stops at the first empty column
- Minimum width: 2 columns

### Height Detection (lines 122-147):

1. **Scans down from header row**: Starting at `start_row + 1`, scans up to **30 rows** below
2. **Stops at empty row**: Checks if a row has any content in the table's column range
3. **Spacer row handling**: Allows one empty row (spacer), but stops at two consecutive empty rows
4. **Content check**: For each row, checks all columns from `start_col` to `end_col` for any non-empty cell

### Boundary Detection Logic:
```python
# Width: Fixed for known tables, or scan up to 10 cols for unknown
# Height: Scan up to 30 rows, stop at first fully empty row (or 2 consecutive empty rows)
max_rows = 30
end_row = start_row
for row in range(start_row + 1, start_row + max_rows):
    # Check if row has content in table columns
    # If empty and next row also empty → stop
```

## 4. Why "Loan to Cost" and "Take Out Loan Sizing" Are Missed

### Possible Reasons:

#### A. Text Mismatch
The most likely cause is that the actual text in Excel doesn't exactly match the search string:
- **"Loan to Cost"** might be stored as:
  - "Loan-to-Cost" (hyphen instead of space)
  - "Loan  to  Cost" (multiple spaces)
  - "Loan to Cost:" (with trailing colon/punctuation)
  - "Loan To Cost" (different capitalization pattern)
  - In a merged cell with different formatting

#### B. Search Finds Wrong Match First
If "Loan to Cost" appears multiple times in the sheet:
- The search finds the **first occurrence** (which might be in a different location)
- If that first match fails boundary expansion, it returns an error
- The code doesn't continue searching for other occurrences

#### C. Header Not in Expected Format
- The header might be split across multiple cells
- The header might be in a merged cell that LibreOffice's search doesn't handle correctly
- The header might be in a different row than expected

#### D. Sheet Search Order
- If the tables are on the same sheet, the search order matters
- "Sources and Uses" (column B) is found because it appears earlier in the sheet
- "Loan to Cost" (column H) and "Take Out Loan Sizing" (column J) appear later, but if the search finds something else first, it stops

### Debugging Steps:

1. **Check stderr logs**: The code logs to stderr (visible in server logs):
   ```
   Searching for 'Loan to Cost' in X sheets
     Found in 'SheetName' at row Y, col Z
   ```

2. **Verify exact text**: Check the exact text in Excel cells H and J:
   - Are there extra spaces?
   - Are there hyphens instead of spaces?
   - Is there trailing punctuation?
   - Is the text in merged cells?

3. **Check if multiple matches exist**: If "Loan to Cost" appears multiple times, the search will only find the first one

## 5. Code Flow Summary

```
find_table_in_all_sheets(doc, "Loan to Cost")
    ↓
For each sheet:
    search = createSearchDescriptor("Loan to Cost", case-insensitive)
    found = findFirst(search)  ← Returns FIRST match only
    ↓
    If found:
        expand_to_table(sheet, found, "Loan to Cost")
            ↓
            start_col = found.Column (e.g., column H = 7)
            start_row = found.Row
            ↓
            Check if "Loan to Cost" in fixed_widths → Yes, width = 8
            end_col = 7 + 8 - 1 = 14
            ↓
            Scan rows down (max 30 rows):
                Check each row for content
                Stop at first fully empty row
            ↓
            Return CellRange(start_col, start_row, end_col, end_row)
        ↓
        Return (sheet, table_range)  ← Stops here, doesn't check other sheets
    ↓
If not found in any sheet:
    Return (None, None)
```

## Recommendations

1. **Add exact text matching verification**: After finding a match, verify the cell text exactly matches the search string
2. **Add fuzzy matching**: Consider using fuzzy string matching for headers
3. **Continue searching on failure**: If boundary expansion fails, continue searching for other occurrences
4. **Add debug output**: Log the exact cell text found to verify matches
5. **Handle merged cells explicitly**: Check if headers are in merged cells and handle accordingly
6. **Add search options**: Consider adding `SearchWords = True` to match whole words only, or `SearchRegularExpression = True` for pattern matching

## Current Fixed Widths

The code has these hardcoded widths:
- `'Sources and Uses'`: 7 columns
- `'Take Out Loan Sizing'`: 3 columns (J, K, L)
- `'Loan to Cost'`: 8 columns
- `'Loan to Value'`: 7 columns
- `'Capital Stack at Closing'`: 7 columns
- `'PILOT Schedule'`: 8 columns

Note: If "Loan to Cost" is found but the width calculation is wrong, it might still fail boundary detection.

