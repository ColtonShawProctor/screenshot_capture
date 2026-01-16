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
KNOWN_TABLE_HEADERS = [
    'sources and uses',
    'take out loan sizing',
    'loan to cost',
    'loan to value',
    'capital stack',
    'release at closing',
    'draw at closing',
    'ltc',
    'ltv',
    'pilot',
    'cost basis',
    'fairbridge metrics',
    'constructional loan',
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
    # Scan right from start_col, stop at 3+ consecutive empty columns with no header styling
    max_cols_to_scan = 50
    end_col = start_col
    consecutive_empty = 0
    max_consecutive_empty = 3
    
    for col in range(start_col, start_col + max_cols_to_scan):
        cell = sheet.getCellByPosition(col, start_row)
        has_content = cell_has_content(cell)
        has_header_style = is_header_cell(cell)
        
        if has_content or has_header_style:
            end_col = col
            consecutive_empty = 0
        else:
            consecutive_empty += 1
            if consecutive_empty >= max_consecutive_empty:
                break
    
    # Also scan down rows to find widest point (data may extend beyond header)
    rows_for_width_detection = 50
    for row in range(start_row, start_row + rows_for_width_detection):
        row_end_col = start_col
        row_consecutive_empty = 0
        
        for col in range(start_col, start_col + max_cols_to_scan):
            cell = sheet.getCellByPosition(col, row)
            if cell_has_content(cell) or is_header_cell(cell):
                row_end_col = col
                row_consecutive_empty = 0
            else:
                row_consecutive_empty += 1
                if row_consecutive_empty >= max_consecutive_empty:
                    break
        
        if row_end_col > end_col:
            end_col = row_end_col
    
    # Ensure minimum width
    if end_col < start_col + 1:
        end_col = start_col + 1
    
    print(f"  Detected width: {end_col - start_col + 1} columns (cols {start_col}-{end_col})", file=sys.stderr)
    
    # Step 2: Determine row boundaries (header-aware)
    max_rows_to_scan = 100
    end_row = start_row
    consecutive_empty_rows = 0
    max_consecutive_empty_rows = 3
    found_total = False
    total_keywords = ['total', 'sum', 'subtotal', 'grand total', 'summary']
    current_table_name_lower = table_name.lower()
    
    for row in range(start_row + 1, start_row + max_rows_to_scan):
        # Check if this row starts a new table (header styling in first column)
        first_cell = sheet.getCellByPosition(start_col, row)
        is_new_header = is_header_cell(first_cell)
        
        if is_new_header:
            # Check if it's a different table's header
            first_cell_text = ""
            try:
                first_cell_text = first_cell.getString().strip().lower()
            except:
                pass
            
            # Check if this header matches a known table header (different from current)
            is_different_table = False
            if first_cell_text:
                for known_header in KNOWN_TABLE_HEADERS:
                    if known_header in first_cell_text and known_header not in current_table_name_lower:
                        # Found a different table's header - stop BEFORE this row
                        print(f"  Stopping at row {row}: detected new table header '{first_cell_text}'", file=sys.stderr)
                        is_different_table = True
                        break
            
            if is_different_table:
                # Stop before this row (don't include it)
                break
            else:
                # Might be a subheader of current table - continue but mark as header row
                print(f"  Row {row}: header styling but same table (subheader?)", file=sys.stderr)
                end_row = row
                consecutive_empty_rows = 0
                continue
        
        # Check if row contains known table header text (even without styling)
        for col in range(start_col, min(end_col + 1, start_col + 3)):
            cell = sheet.getCellByPosition(col, row)
            try:
                cell_text = cell.getString().strip().lower()
                for known_header in KNOWN_TABLE_HEADERS:
                    if cell_text == known_header or cell_text.startswith(known_header):
                        if known_header not in current_table_name_lower:
                            # Found a different table's header - stop BEFORE
                            print(f"  Stopping at row {row}: found different table header '{cell_text}'", file=sys.stderr)
                            return sheet.getCellRangeByPosition(start_col, start_row, end_col, row - 1)
            except:
                pass
        
        # Check if row is completely empty
        row_empty = True
        row_text = ""
        for col in range(start_col, end_col + 1):
            cell = sheet.getCellByPosition(col, row)
            if cell_has_content(cell):
                row_empty = False
                try:
                    cell_text = cell.getString().strip().lower()
                    if cell_text:
                        row_text += " " + cell_text
                except:
                    pass
        
        if row_empty:
            consecutive_empty_rows += 1
            if consecutive_empty_rows >= max_consecutive_empty_rows:
                # Found multiple consecutive empty rows - table has ended
                print(f"  Stopping at row {row}: {consecutive_empty_rows} consecutive empty rows", file=sys.stderr)
                break
        else:
            consecutive_empty_rows = 0
            end_row = row
            
            # Check for "Total" row - include it then stop after footnotes
            first_cell_text = ""
            try:
                first_cell_text = first_cell.getString().strip().lower()
            except:
                pass
            
            if first_cell_text.startswith('total') and not found_total:
                found_total = True
                print(f"  Found total row at {row}, scanning for footnotes...", file=sys.stderr)
                # Continue for 1-2 more rows to catch footnotes (lines starting with *)
                for extra_row in range(row + 1, min(row + 3, start_row + max_rows_to_scan)):
                    extra_empty = True
                    for col in range(start_col, end_col + 1):
                        extra_cell = sheet.getCellByPosition(col, extra_row)
                        if cell_has_content(extra_cell):
                            extra_cell_text = ""
                            try:
                                extra_cell_text = extra_cell.getString().strip()
                            except:
                                pass
                            # Include footnote rows (start with *) or empty rows after total
                            if not extra_cell_text or extra_cell_text.startswith('*'):
                                extra_empty = False
                                break
                            else:
                                # Non-footnote content - might be next table starting
                                extra_empty = True
                                break
                    
                    if not extra_empty:
                        end_row = extra_row
                    else:
                        break
                # Stop after total + footnotes
                break
    
    # Add small buffer for borders/formatting
    if end_row < start_row + max_rows_to_scan - 1:
        end_row += 1
    
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