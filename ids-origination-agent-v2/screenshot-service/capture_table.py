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


def expand_to_table(sheet, header_cell, table_name):
    """
    Dynamically find table boundaries by scanning for content.
    Detects columns until multiple consecutive empty columns.
    Detects rows until multiple consecutive empty rows, ensuring totals are included.
    """
    start_col = header_cell.CellAddress.Column
    start_row = header_cell.CellAddress.Row
    
    print(f"  Expanding '{table_name}' from row {start_row}, col {start_col}", file=sys.stderr)
    
    # First, scan header row to get initial column width estimate
    # This helps us know what columns to check for row detection
    max_cols_to_scan = 50  # Increased from 10
    initial_end_col = start_col
    
    # Scan header row to find where it ends
    for col in range(start_col, start_col + max_cols_to_scan):
        cell = sheet.getCellByPosition(col, start_row)
        if cell_has_content(cell):
            initial_end_col = col
        else:
            # Allow one empty column, but if next is also empty, stop
            if col + 1 < start_col + max_cols_to_scan:
                next_cell = sheet.getCellByPosition(col + 1, start_row)
                if not cell_has_content(next_cell):
                    break
    
    # Now scan down rows to find the actual table width
    # Tables may have data columns beyond the header row
    # Scan first 50 rows to find the widest point
    rows_for_width_detection = 50
    end_col = initial_end_col
    
    # Scan multiple rows to find the widest point of the table
    for row in range(start_row, start_row + rows_for_width_detection):
        # Scan right from start_col to find content in this row
        row_end_col = start_col
        consecutive_empty = 0
        max_consecutive_empty = 3  # Stop after 3 consecutive empty columns
        
        for col in range(start_col, start_col + max_cols_to_scan):
            cell = sheet.getCellByPosition(col, row)
            if cell_has_content(cell):
                row_end_col = col
                consecutive_empty = 0
            else:
                consecutive_empty += 1
                if consecutive_empty >= max_consecutive_empty:
                    break
        
        # Update end_col if this row extends further
        if row_end_col > end_col:
            end_col = row_end_col
    
    # Ensure minimum width
    if end_col < start_col + 1:
        end_col = start_col + 1
    
    print(f"  Detected width: {end_col - start_col + 1} columns (cols {start_col}-{end_col})", file=sys.stderr)
    
    # Now detect row height - scan down until multiple consecutive empty rows
    max_rows_to_scan = 100  # Increased limit
    end_row = start_row
    consecutive_empty_rows = 0
    max_consecutive_empty_rows = 3  # Stop after 3 consecutive empty rows
    
    # Keywords that indicate summary/total rows - ensure we include them
    total_keywords = ['total', 'sum', 'subtotal', 'grand total', 'summary']
    
    for row in range(start_row + 1, start_row + max_rows_to_scan):
        row_has_content = False
        row_text = ""
        
        # Check all columns in the detected table width
        for col in range(start_col, end_col + 1):
            cell = sheet.getCellByPosition(col, row)
            if cell_has_content(cell):
                row_has_content = True
                # Collect text for keyword checking
                cell_text = cell.getString().strip().lower()
                if cell_text:
                    row_text += " " + cell_text
        
        if row_has_content:
            end_row = row
            consecutive_empty_rows = 0
            
            # Check if this row contains total keywords - if so, scan a bit further
            # to catch any additional summary rows
            if any(keyword in row_text for keyword in total_keywords):
                # Found a total row - scan 2-3 more rows to catch additional totals
                for lookahead_row in range(row + 1, min(row + 4, start_row + max_rows_to_scan)):
                    lookahead_has_content = False
                    for col in range(start_col, end_col + 1):
                        lookahead_cell = sheet.getCellByPosition(col, lookahead_row)
                        if cell_has_content(lookahead_cell):
                            lookahead_has_content = True
                            break
                    if lookahead_has_content:
                        end_row = lookahead_row
                    else:
                        break
        else:
            consecutive_empty_rows += 1
            if consecutive_empty_rows >= max_consecutive_empty_rows:
                # Found multiple consecutive empty rows - table has ended
                break
    
    # Add a small buffer (1 row) to ensure we don't cut off borders/formatting
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