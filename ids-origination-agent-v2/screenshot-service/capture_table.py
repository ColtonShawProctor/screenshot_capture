#!/usr/bin/env python3
"""
Capture Excel table as image using LibreOffice UNO API.
Searches ALL sheets for the header text, then exports that range as PNG.
"""

import sys
import json
import base64
import tempfile
import os
import time
import subprocess

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.table import CellRangeAddress


def connect_to_libreoffice(max_retries=5):
    """Connect to running LibreOffice instance."""
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context
    )
    
    for attempt in range(max_retries):
        try:
            ctx = resolver.resolve(
                "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
            )
            smgr = ctx.ServiceManager
            desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
            return desktop
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(1)
            else:
                raise RuntimeError(f"Cannot connect to LibreOffice: {e}")


def clear_all_print_areas(doc):
    """Clear print areas from ALL sheets to prevent accumulation."""
    sheets = doc.getSheets()
    for i in range(sheets.getCount()):
        sheet = sheets.getByIndex(i)
        sheet.setPrintAreas(())  # Empty tuple clears


def find_table_in_all_sheets(doc, header_text):
    """
    Search ALL sheets for the header text with debug logging.
    Returns (sheet, cell_range) or (None, None) if not found.
    """
    sheets = doc.getSheets()
    
    print(f"Searching for '{header_text}' in {sheets.getCount()} sheets", file=sys.stderr)
    
    for i in range(sheets.getCount()):
        sheet = sheets.getByIndex(i)
        sheet_name = sheet.getName()
        
        # Create search descriptor
        search = sheet.createSearchDescriptor()
        search.SearchString = header_text
        search.SearchCaseSensitive = False
        
        found = sheet.findFirst(search)
        if found:
            cell_addr = found.getCellAddress()
            print(f"  Found in '{sheet_name}' at row {cell_addr.Row}, col {cell_addr.Column}", file=sys.stderr)
            
            # Found the header - now expand to get full table
            table_range = expand_to_table(sheet, found, header_text)
            if table_range:
                return sheet, table_range
            else:
                print(f"  Could not expand to valid table range in '{sheet_name}'", file=sys.stderr)
        else:
            print(f"  Not found in '{sheet_name}'", file=sys.stderr)
    
    return None, None


def cell_has_content(cell):
    """Check if a cell has any content (text, numbers, formulas)."""
    try:
        # Check for string content
        if cell.getString().strip():
            return True
        # Check for numeric value
        if cell.getValue() is not None:
            val = cell.getValue()
            if isinstance(val, (int, float)) and val != 0:
                return True
        # Check for formula
        if cell.getFormula():
            return True
    except:
        pass
    return False


# Known table headers to detect adjacent tables
# IMPORTANT: Use EXACT or near-exact matches only
# Don't include short strings like 'ltv' or 'ltc' alone - they appear as row labels within other tables
# 
# DO NOT ADD:
# - "constructional loan" - it's a COLUMN HEADER within Sources and Uses
# - "value stress test" - it's a SUBHEADER within LTV Sensitivity Table
# - "takeout financing chart" - it's a SUBHEADER within LTV Sensitivity Table
#
KNOWN_TABLE_HEADERS = [
    'sources and uses',
    'take out loan sizing',
    'loan to cost',       # Full name only - not just 'ltc'
    'loan to value',      # Full name only - not just 'ltv'
    'capital stack at closing',
    'capital stack',
    'release at closing',
    'draw at closing',    # This starts the Draw at Closing mini-table
    'pilot schedule',
    'cost basis',
    'fairbridge metrics',
    'disbursements at closing',
    'ltc',                # Standalone 'LTC' header (not 'LTC (At Closing)' which is a row label)
]

# Headers that indicate a NEW table even if they share words with current table
# "Capital Stack at Closing" has "Sources" column which matches "Sources and Uses"
STRONG_TABLE_HEADERS = [
    'capital stack at closing',
    'capital stack',
    'draw at closing',
    'ltc',
    'loan to cost',
    'loan to value',
]


def is_header_cell(cell):
    """
    Check if cell has dark header styling (navy background).
    Fairbridge headers use dark blue: RGB approximately (0, 32, 96) or similar.
    LibreOffice UNO: CellBackColor is a 32-bit integer in format 0xAARRGGBB
    """
    try:
        # Get cell background color property
        # LibreOffice uses getPropertyValue for cell properties
        props = cell.getPropertySetInfo()
        if props.hasPropertyByName("CellBackColor"):
            color = cell.getPropertyValue("CellBackColor")
            if color is not None and isinstance(color, int):
                # Extract RGB components from 32-bit integer (0xAARRGGBB format)
                # Alpha is high byte, but we check RGB
                r = (color >> 16) & 0xFF
                g = (color >> 8) & 0xFF
                b = color & 0xFF
                
                # Dark blue detection: low R, low-medium G, higher B
                # Common Fairbridge header colors: RGB(0,32,96), RGB(15,36,62), RGB(0,51,102)
                if r < 50 and g < 80 and b > 50:
                    print(f"    Header cell detected: RGB({r},{g},{b})", file=sys.stderr)
                    return True
    except Exception as e:
        # If property access fails, fall back to content-based detection
        pass
    return False


def is_different_table_header(cell_text, current_table_name):
    """
    Check if cell_text indicates a DIFFERENT table is starting.
    
    CRITICAL CASES:
    - 'LTV' row inside Take Out Loan Sizing -> NOT a new table (it's a metric row)
    - 'Draw at Closing' header -> IS a new table
    - 'LTC' header -> IS a new table (but 'LTC (At Closing)' is a row label)
    """
    cell_text = cell_text.lower().strip()
    current_lower = current_table_name.lower()
    
    # Skip if empty
    if not cell_text:
        return False
    
    # Skip if this text is part of our current table name
    if cell_text in current_lower or current_lower in cell_text:
        return False
    
    # Special handling for short labels that could be row labels OR table headers
    # 'ltc' alone can be a table header (standalone LTC table)
    # 'ltc (at closing)' is a row label, not a table header
    # 'ltv' alone is NOT a table header - it's often a row label (e.g., in Take Out Loan Sizing)
    # Only "loan to value" (full name) is a table header
    if cell_text == 'ltc':
        return True  # Standalone 'ltc' = table header
    if cell_text.startswith('ltc '):
        return False  # Has suffix = row label like "LTC (At Closing)"
    # Don't treat standalone 'ltv' as table header - it's usually a row label
    
    for known_header in KNOWN_TABLE_HEADERS:
        # Exact match
        if cell_text == known_header:
            return True
        # Cell starts with known header (e.g., "Draw at Closing" matches "draw at closing")
        if cell_text.startswith(known_header) and len(cell_text) <= len(known_header) + 10:
            return True
    
    return False


def find_column_boundaries(sheet, header_row, start_col):
    """
    Scan right from start_col to find where table ends horizontally.
    
    CRITICAL: Tables often have empty spacing columns between label and value sections!
    Example - Loan to Cost table layout:
    Col 7: "Loan to Cost" (header)
    Col 8: (empty spacing)
    Col 9: "At-Closing" values
    Col 10: (empty spacing)  
    Col 11: "W/ Carry Costs" values
    
    Must use 4+ consecutive empty columns as threshold, not 2.
    Small gaps (1-3 cols) are spacing within the same table.
    """
    max_col = start_col
    consecutive_empty = 0
    max_cols_to_scan = 100  # Reasonable maximum for Excel sheets
    
    # Scan header row first
    for col in range(start_col, start_col + max_cols_to_scan):
        try:
            cell = sheet.getCellByPosition(col, header_row)
            has_content = cell_has_content(cell)
            
            # IMPORTANT: Only extend if this cell has content AND is adjacent to previous content
            # Don't extend just because a cell has header styling - it could be a different table
            if has_content:
                # Check if we skipped empty columns to get here
                if consecutive_empty >= 4:
                    # There was a LARGE gap (4+ empty cols) - this is a different table, stop
                    # Small gaps (1-3 cols) are spacing within the same table
                    break
                max_col = col
                consecutive_empty = 0
            else:
                consecutive_empty += 1
                # Use 4+ consecutive empty columns as threshold (allows 1-3 col spacing within tables)
                if consecutive_empty >= 4:
                    break
        except:
            # Out of bounds - stop scanning
            break
    
    # Also scan down a few rows to find widest point (data may extend beyond header)
    rows_to_scan = 50
    for row in range(header_row, header_row + rows_to_scan):
        row_max_col = start_col
        row_consecutive_empty = 0
        
        for col in range(start_col, start_col + max_cols_to_scan):
            try:
                cell = sheet.getCellByPosition(col, row)
                has_content = cell_has_content(cell)
                
                if has_content:
                    if row_consecutive_empty >= 4:
                        break
                    row_max_col = col
                    row_consecutive_empty = 0
                else:
                    row_consecutive_empty += 1
                    # Use 4+ consecutive empty columns as threshold (allows 1-3 col spacing within tables)
                    if row_consecutive_empty >= 4:
                        break
            except:
                # Out of bounds - stop scanning this row
                break
        
        if row_max_col > max_col:
            max_col = row_max_col
    
    # Add 1 column buffer for any overflow
    return max_col + 1


def find_row_boundaries(sheet, header_row, start_col, end_col, current_table_name):
    """
    Scan down from header_row to find where table ends vertically.
    
    STOP BEFORE a row if:
    - ANY cell in the row has header-style background AND contains a known table name
    - 3+ consecutive completely empty rows (after checking for Total ahead)
    
    CRITICAL FIX: Check ALL columns for header cells, not just start_col!
    "Draw at Closing" header might be in column 10 while we started at column 9.
    
    IMPORTANT: Do NOT stop at "Total" rows - some tables have content after Total!
    Example: Release at Closing has:
      - Total Disbursements (row 16)
      - (-) Sponsor's Equity at Closing (row 18)  <- MUST INCLUDE
      - Fairbridge Release at Closing (row 19)   <- MUST INCLUDE
    """
    max_row = header_row
    consecutive_empty = 0
    max_rows_to_scan = 200  # Reasonable maximum for Excel sheets
    
    for row in range(header_row + 1, header_row + max_rows_to_scan):
        try:
            # CRITICAL: Check ALL columns in this row for header styling
            found_new_table = False
            for col in range(start_col, end_col + 1):
                try:
                    cell = sheet.getCellByPosition(col, row)
                    cell_text = ""
                    try:
                        cell_text = cell.getString().strip()
                    except:
                        pass
                    
                    if is_header_cell(cell) and cell_text:
                        cell_lower = cell_text.lower()
                        print(f"    Header cell detected at row {row}, col {col}: '{cell_text}'", file=sys.stderr)
                        
                        # Check if this is a STRONG table header (always stops, even if name overlaps)
                        for strong_header in STRONG_TABLE_HEADERS:
                            if cell_lower == strong_header or cell_lower.startswith(strong_header):
                                # But skip if it's our own table
                                if strong_header not in current_table_name.lower():
                                    print(f"  Stopping at row {row}: detected strong table header '{cell_lower}'", file=sys.stderr)
                                    found_new_table = True
                                    break
                        
                        if found_new_table:
                            break
                            
                        # Regular check for different table
                        if is_different_table_header(cell_text, current_table_name):
                            print(f"  Stopping at row {row}: detected new table header '{cell_lower}'", file=sys.stderr)
                            found_new_table = True
                            break
                        else:
                            print(f"  Row {row}: header styling but same table (subheader?) text='{cell_text}'", file=sys.stderr)
                except:
                    # Out of bounds - continue to next column
                    pass
            
            if found_new_table:
                break
            
            # Second pass: Check for STRONG table names even WITHOUT blue styling
            # This catches headers like "Draw at Closing" that might have different styling
            for col in range(start_col, end_col + 1):
                try:
                    cell = sheet.getCellByPosition(col, row)
                    cell_text = ""
                    try:
                        cell_text = cell.getString().strip()
                    except:
                        pass
                    cell_text_lower = cell_text.lower() if cell_text else ""
                    
                    if cell_text_lower:
                        # Check for STRONG table headers even without blue styling
                        for strong_header in STRONG_TABLE_HEADERS:
                            if cell_text_lower == strong_header or cell_text_lower.startswith(strong_header + ' '):
                                # But skip if it's our own table
                                if strong_header not in current_table_name.lower():
                                    print(f"  Stopping at row {row}: found strong table name '{cell_text_lower}' (no blue styling)", file=sys.stderr)
                                    found_new_table = True
                                    break
                        if found_new_table:
                            break
                except:
                    # Out of bounds - continue to next column
                    pass
            
            if found_new_table:
                break
            
            # Check the first cell for logging
            first_cell = sheet.getCellByPosition(start_col, row)
            first_cell_text = ""
            try:
                first_cell_text = first_cell.getString().strip()
            except:
                pass
            first_cell_lower = first_cell_text.lower()
            
            # Check if row is completely empty
            row_empty = True
            for col in range(start_col, end_col + 1):
                try:
                    cell = sheet.getCellByPosition(col, row)
                    if cell_has_content(cell):
                        row_empty = False
                        break
                except:
                    # Out of bounds - treat as empty
                    pass
            
            if row_empty:
                consecutive_empty += 1
                if consecutive_empty >= 3:
                    # Before stopping, scan ahead 5 rows to check for a Total row
                    # Some tables have empty spacing before the Total
                    # Check ALL columns in the scanned rows, not just start_col
                    found_total_ahead = False
                    for scan_row in range(row + 1, min(row + 6, header_row + max_rows_to_scan)):
                        try:
                            # Check all columns in this row for "Total"
                            for scan_col in range(start_col, end_col + 1):
                                try:
                                    scan_cell = sheet.getCellByPosition(scan_col, scan_row)
                                    scan_text = ""
                                    try:
                                        scan_text = scan_cell.getString().strip()
                                    except:
                                        pass
                                    scan_text_lower = scan_text.lower().strip()
                                    if scan_text_lower.startswith('total'):
                                        print(f"  Found Total row at {scan_row} after empty gap, extending...", file=sys.stderr)
                                        found_total_ahead = True
                                        max_row = scan_row
                                        consecutive_empty = 0
                                        break
                                except:
                                    pass
                            if found_total_ahead:
                                break
                        except:
                            pass
                    
                    if not found_total_ahead:
                        print(f"  Stopping at row {row}: 3 consecutive empty rows", file=sys.stderr)
                        break
            else:
                consecutive_empty = 0
                max_row = row
                
                # Log total rows for debugging, but DON'T stop
                if first_cell_lower.startswith('total'):
                    print(f"  Found total row at {row}, continuing scan...", file=sys.stderr)
        except:
            # Out of bounds - stop scanning
            break
    
    return max_row + 1  # +1 buffer


def expand_to_table(sheet, header_cell, table_name):
    """
    Header-based table boundary detection.
    Uses dark blue header bars as definitive table boundaries.
    Detects columns until multiple consecutive empty columns (with no header styling).
    Detects rows until hitting a new table header or 3+ consecutive empty rows.
    """
    start_col = header_cell.CellAddress.Column
    start_row = header_cell.CellAddress.Row
    
    print(f"  Expanding '{table_name}' from row {start_row}, col {start_col}", file=sys.stderr)
    
    # Step 1: Determine column boundaries
    end_col = find_column_boundaries(sheet, start_row, start_col)
    
    # Ensure minimum width
    if end_col < start_col + 1:
        end_col = start_col + 1
    
    print(f"  Detected width: {end_col - start_col + 1} columns (cols {start_col}-{end_col})", file=sys.stderr)
    
    # Step 2: Determine row boundaries
    end_row = find_row_boundaries(sheet, start_row, start_col, end_col, table_name)
    
    print(f"  Detected height: {end_row - start_row + 1} rows (rows {start_row}-{end_row})", file=sys.stderr)
    print(f"  Final range: cols {start_col}-{end_col}, rows {start_row}-{end_row}", file=sys.stderr)
    
    return sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)


def export_range_as_image(doc, sheet, table_range, output_path):
    """
    Set print area to table range, export as PDF, convert to PNG.
    """
    # Activate the sheet
    controller = doc.getCurrentController()
    controller.setActiveSheet(sheet)
    
    # Force recalculation
    doc.calculateAll()
    
    # Get the range address for print area
    range_address = table_range.getRangeAddress()
    
    # Set print area to ONLY this range
    # Clear any existing print areas first
    sheet.setPrintAreas(())
    
    # Create a new print area with just our range
    print_area = CellRangeAddress()
    print_area.Sheet = range_address.Sheet
    print_area.StartColumn = range_address.StartColumn
    print_area.StartRow = range_address.StartRow
    print_area.EndColumn = range_address.EndColumn
    print_area.EndRow = range_address.EndRow
    
    sheet.setPrintAreas((print_area,))
    
    # Configure page style for minimal margins
    style_families = doc.getStyleFamilies()
    page_styles = style_families.getByName("PageStyles")
    default_style = page_styles.getByName("Default")
    
    # Set minimal margins (in 1/100 mm)
    default_style.setPropertyValue("LeftMargin", 500)
    default_style.setPropertyValue("RightMargin", 500)
    default_style.setPropertyValue("TopMargin", 500)
    default_style.setPropertyValue("BottomMargin", 500)
    
    # Export as PDF
    pdf_path = output_path.replace('.png', '.pdf')
    pdf_url = uno.systemPathToFileUrl(pdf_path)
    
    export_props = (
        PropertyValue(Name="FilterName", Value="calc_pdf_Export"),
    )
    
    doc.storeToURL(pdf_url, export_props)
    
    # Convert PDF to PNG using pdftoppm
    png_base = output_path.replace('.png', '')
    subprocess.run([
        'pdftoppm', '-png', '-r', '150', '-singlefile',
        pdf_path, png_base
    ], check=True, capture_output=True)
    
    # pdftoppm outputs to {base}.png
    actual_png = png_base + '.png'
    if actual_png != output_path:
        os.rename(actual_png, output_path)
    
    # Trim whitespace using ImageMagick
    subprocess.run([
        'convert', output_path, '-trim', '+repage', output_path
    ], check=True, capture_output=True)
    
    # Cleanup PDF
    if os.path.exists(pdf_path):
        os.unlink(pdf_path)


def capture_single_table(desktop, excel_path, table_name, output_path):
    """Open doc, capture ONE table, close doc."""
    
    # Open fresh document
    file_url = uno.systemPathToFileUrl(excel_path)
    load_props = (
        PropertyValue(Name="Hidden", Value=True),
        PropertyValue(Name="ReadOnly", Value=True),
    )
    doc = desktop.loadComponentFromURL(file_url, "_blank", 0, load_props)
    
    if not doc:
        raise RuntimeError("Failed to open document")
    
    try:
        # Clear ALL print areas on ALL sheets first
        clear_all_print_areas(doc)
        
        # Find and capture the table
        sheet, table_range = find_table_in_all_sheets(doc, table_name)
        
        if not sheet or not table_range:
            return False, f"Table '{table_name}' not found in any sheet"
        
        # Export the range
        export_range_as_image(doc, sheet, table_range, output_path)
        
        return True, "Success"
        
    finally:
        doc.close(True)  # ALWAYS close


def main():
    """Main entry point."""
    try:
        input_data = json.load(sys.stdin)
        
        excel_base64 = input_data['excelBase64']
        table_name = input_data['tableName']
        
        # Save Excel to temp file
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
            f.write(base64.b64decode(excel_base64))
            excel_path = f.name
        
        output_path = tempfile.mktemp(suffix='.png')
        
        try:
            # Connect to LibreOffice
            desktop = connect_to_libreoffice()
            
            # Capture single table with fresh document
            success, message = capture_single_table(desktop, excel_path, table_name, output_path)
            
            if not success:
                print(json.dumps({
                    'success': False,
                    'error': message
                }))
                return
            
            # Read result
            with open(output_path, 'rb') as f:
                image_base64 = base64.b64encode(f.read()).decode('utf-8')
            
            print(json.dumps({
                'success': True,
                'image': image_base64
            }))
            
        finally:
            # Cleanup
            if os.path.exists(excel_path):
                os.unlink(excel_path)
            if os.path.exists(output_path):
                os.unlink(output_path)
                
    except Exception as e:
        print(json.dumps({
            'success': False,
            'error': str(e)
        }))
        sys.exit(1)


if __name__ == '__main__':
    main()