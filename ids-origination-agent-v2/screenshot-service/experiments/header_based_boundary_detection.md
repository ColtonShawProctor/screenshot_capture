# Experiment: Header-Based Table Boundary Detection

## Goal

**Impact**: Improve table screenshot accuracy by using visual header markers (dark blue header bars) as definitive table boundaries. This will:
- Prevent capturing multiple adjacent tables in one screenshot
- Ensure totals and footnotes are included
- Work reliably across different Excel layouts without hardcoded widths
- Reduce false positives from empty row/column gaps

**Current Problem**: The existing algorithm uses content-based detection (empty rows/columns) which fails when:
- Tables are close together (1-2 row gaps)
- Tables are side-by-side
- Subheaders exist within a table
- Footnotes follow totals

## What I Learned from Existing Code

### Current Implementation (`expand_to_table` function)
1. **Width Detection**: 
   - Scans up to 50 columns from header
   - Finds widest point across first 50 rows
   - Stops at 3 consecutive empty columns
   - No header styling detection

2. **Height Detection**:
   - Scans up to 100 rows
   - Stops at 3 consecutive empty rows
   - Has total keyword detection but continues scanning
   - Adds 1-row buffer

3. **Limitations**:
   - No visual header detection (relies on content gaps)
   - Can't distinguish between subheaders and new table headers
   - Doesn't stop before adjacent tables reliably
   - Fixed width dictionary was removed in favor of dynamic scanning

### LibreOffice UNO API Capabilities
- `sheet.getCellByPosition(col, row)` - Get cell
- `cell.getCellPropertySet()` - Access cell properties
- `cell.getPropertyValue("CellBackColor")` - Get background color (returns long integer)
- `cell.getString()`, `cell.getValue()`, `cell.getFormula()` - Content access
- Color format: LibreOffice uses 32-bit integer (0xAARRGGBB format)

### Key Insight
Fairbridge Excel models use consistent dark blue headers (RGB ~0,32,96 or similar). These are visual markers that can be detected programmatically via cell background color, providing a more reliable boundary than content gaps.

## Surgical Plan

### Phase 1: Add Header Detection Function
1. Create `is_header_cell(cell)` function that:
   - Gets cell background color via UNO API
   - Checks if RGB values indicate dark blue (low R, low-medium G, higher B)
   - Handles LibreOffice color format (32-bit integer)

### Phase 2: Enhance Column Detection
1. Modify column boundary detection to:
   - Check for header styling in addition to content
   - Stop when hitting 3+ consecutive columns with no content AND no header styling
   - This handles side-by-side tables better

### Phase 3: Rewrite Row Detection (Critical)
1. Replace current row scanning logic with header-aware detection:
   - Scan down from header row
   - **STOP BEFORE** any row that has header styling AND contains a known table header text (different from current table)
   - **STOP AFTER** "Total" row (include it, then check for footnotes)
   - Handle subheaders: if header styling but same table name or no known table name, continue
   - Still use 3 consecutive empty rows as fallback

### Phase 4: Known Table Headers List
1. Add list of known table headers to detect adjacent tables:
   - Sources and Uses, Take Out Loan Sizing, Loan to Cost, Loan to Value, Capital Stack, etc.

### Phase 5: Testing & Debugging
1. Add verbose logging for:
   - Header cell detection (RGB values)
   - Row-by-row decisions (why we stop/continue)
   - Column boundary decisions

## Implementation Notes

- LibreOffice UNO color format: `getPropertyValue("CellBackColor")` returns a long integer
- Need to convert to RGB: `(color >> 16) & 0xFF` for red, `(color >> 8) & 0xFF` for green, `color & 0xFF` for blue
- Alpha channel: `(color >> 24) & 0xFF` (usually 0 for opaque)
- Dark blue detection: R < 50, G < 80, B > 50 (approximate)

---

## Attempted Solution

### Implementation Summary

**Added Functions:**
1. `is_header_cell(cell)` - Detects dark blue header cells by checking `CellBackColor` property
   - Extracts RGB from 32-bit integer (0xAARRGGBB format)
   - Returns True if R < 50, G < 80, B > 50 (dark blue range)
   - Includes error handling for property access failures

2. `KNOWN_TABLE_HEADERS` - List of known table header strings to detect adjacent tables

**Modified `expand_to_table()` Function:**

**Column Detection (Enhanced):**
- Scans right from start_col, checking both content AND header styling
- Stops at 3+ consecutive empty columns with no header styling
- Also scans down first 50 rows to find widest point (data may extend beyond header)

**Row Detection (Rewritten - Header-Aware):**
- **Primary Stop Condition**: Detects new table headers by:
  - Checking if first cell in row has header styling (`is_header_cell()`)
  - If header styling found, checks if cell text matches a known table header
  - If it's a DIFFERENT table (not current table name), stops BEFORE that row
  - If it's same table (subheader), continues scanning
  
- **Secondary Stop Conditions**:
  - Checks for known table header text in first 3 columns (even without styling)
  - Stops at 3+ consecutive completely empty rows (fallback)
  
- **Total Row Handling**:
  - Detects rows starting with "total" keyword
  - Includes total row, then scans 1-2 more rows for footnotes (lines starting with *)
  - Stops after total + footnotes

- **Subheader Handling**:
  - If header styling found but text matches current table or no known table, treats as subheader
  - Continues scanning (doesn't stop)

**Debug Logging:**
- Logs RGB values when header cells detected
- Logs reason for stopping (new table header, empty rows, etc.)
- Logs total row detection

### Key Changes from Original:
1. **Header Detection**: Now uses visual styling (background color) instead of just content gaps
2. **Stop Before New Tables**: Actively detects and stops before adjacent table headers
3. **Subheader Awareness**: Distinguishes between subheaders (same table) and new table headers
4. **Total + Footnotes**: Explicitly handles total rows and footnote rows
5. **Known Headers List**: Uses list of known table names to identify adjacent tables

### Technical Details:
- LibreOffice UNO API: Uses `cell.getPropertyValue("CellBackColor")` to access background color
- Color format: 32-bit integer in 0xAARRGGBB format
- RGB extraction: `r = (color >> 16) & 0xFF`, `g = (color >> 8) & 0xFF`, `b = color & 0xFF`
- Error handling: Falls back gracefully if property access fails

### Files Modified:
- `capture_table.py`: Added `is_header_cell()`, `KNOWN_TABLE_HEADERS`, rewrote `expand_to_table()`

