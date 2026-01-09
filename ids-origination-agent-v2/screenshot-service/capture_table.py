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
            table_range = expand_to_table(sheet, found)
            if table_range:
                return sheet, table_range
            else:
                print(f"  Could not expand to valid table range in '{sheet_name}'", file=sys.stderr)
        else:
            print(f"  Not found in '{sheet_name}'", file=sys.stderr)
    
    return None, None


def expand_to_table(sheet, header_cell):
    """
    Starting from header cell, expand to find table boundaries.
    Stop when we hit truly empty rows/columns.
    """
    start_col = header_cell.CellAddress.Column
    start_row = header_cell.CellAddress.Row
    
    print(f"  Expanding table from row {start_row}, col {start_col}", file=sys.stderr)
    
    # Find table width - scan right until 2+ consecutive empty cells in header row
    end_col = start_col
    empty_count = 0
    for col in range(start_col, start_col + 30):  # Max 30 cols
        cell = sheet.getCellByPosition(col, start_row)
        has_content = cell.getString().strip() != '' or cell.getType() != 0
        
        if has_content:
            end_col = col
            empty_count = 0
        else:
            empty_count += 1
            if empty_count >= 2:  # 2 consecutive empty = end of table
                break
    
    # Find table height - scan down until 2+ consecutive fully empty rows
    end_row = start_row
    empty_rows = 0
    for row in range(start_row + 1, start_row + 50):  # Max 50 rows
        # Check if ANY cell in this row (within table width) has content
        row_has_content = False
        for col in range(start_col, end_col + 1):
            cell = sheet.getCellByPosition(col, row)
            if cell.getString().strip() != '' or cell.getType() != 0:
                row_has_content = True
                break
        
        if row_has_content:
            end_row = row
            empty_rows = 0
        else:
            empty_rows += 1
            if empty_rows >= 2:  # 2 consecutive empty rows = end of table
                break
    
    print(f"  Table range: cols {start_col}-{end_col}, rows {start_row}-{end_row}", file=sys.stderr)
    
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