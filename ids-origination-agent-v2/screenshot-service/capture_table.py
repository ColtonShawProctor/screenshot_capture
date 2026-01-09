#!/usr/bin/env python3
"""
Capture Excel table as image using LibreOffice UNO API.
Finds table by header text, selects the range, exports as PNG.
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


def find_table_range(sheet, header_text):
    """
    Find a table by its header text and return the cell range.
    Uses LibreOffice's search and expands to find table boundaries.
    """
    # Create search descriptor
    search = sheet.createSearchDescriptor()
    search.SearchString = header_text
    search.SearchCaseSensitive = False
    
    found = sheet.findFirst(search)
    if not found:
        return None
    
    # Get the starting cell position
    start_col = found.CellAddress.Column
    start_row = found.CellAddress.Row
    
    # Expand down to find table end (stop at empty row)
    end_row = start_row
    end_col = start_col
    
    # First, find the width by scanning the header row
    cursor = sheet.createCursor()
    cursor.gotoCell(sheet.getCellByPosition(start_col, start_row), False)
    
    # Scan right in header row to find table width
    col = start_col
    while True:
        cell = sheet.getCellByPosition(col, start_row)
        if cell.getType() == 0 and col > start_col:  # EMPTY and not first col
            # Check if next few columns are also empty (handle gaps)
            all_empty = True
            for check_col in range(col, min(col + 3, 100)):
                if sheet.getCellByPosition(check_col, start_row).getType() != 0:
                    all_empty = False
                    break
            if all_empty:
                break
        end_col = col
        col += 1
        if col > 100:  # Safety limit
            break
    
    # Scan down to find table height
    row = start_row
    consecutive_empty = 0
    while consecutive_empty < 2:  # Allow 1 empty row within table
        row += 1
        if row > 500:  # Safety limit
            break
        
        # Check if entire row in table range is empty
        row_empty = True
        for c in range(start_col, end_col + 1):
            cell = sheet.getCellByPosition(c, row)
            if cell.getType() != 0 or cell.getString().strip():
                row_empty = False
                break
        
        if row_empty:
            consecutive_empty += 1
        else:
            consecutive_empty = 0
            end_row = row
    
    # Return the range
    return sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)


def export_range_as_image(desktop, excel_path, sheet_name, header_text, output_path):
    """
    Open Excel, find table by header, export as PNG.
    """
    # Load the document (hidden)
    file_url = uno.systemPathToFileUrl(excel_path)
    load_props = (
        PropertyValue(Name="Hidden", Value=True),
        PropertyValue(Name="ReadOnly", Value=True),
    )
    
    doc = desktop.loadComponentFromURL(file_url, "_blank", 0, load_props)
    
    if not doc:
        raise RuntimeError("Failed to open document")
    
    try:
        # Get the sheet
        sheets = doc.getSheets()
        sheet = None
        
        # Try exact match first
        if sheets.hasByName(sheet_name):
            sheet = sheets.getByName(sheet_name)
        else:
            # Try with/without trailing space
            for i in range(sheets.getCount()):
                s = sheets.getByIndex(i)
                if s.getName().strip() == sheet_name.strip():
                    sheet = s
                    break
        
        if not sheet:
            # If sheet not found, search all sheets for the header
            for i in range(sheets.getCount()):
                s = sheets.getByIndex(i)
                test_range = find_table_range(s, header_text)
                if test_range:
                    sheet = s
                    break
        
        if not sheet:
            raise RuntimeError(f"Sheet '{sheet_name}' not found")
        
        # Activate the sheet
        controller = doc.getCurrentController()
        controller.setActiveSheet(sheet)
        
        # Find the table
        table_range = find_table_range(sheet, header_text)
        if not table_range:
            raise RuntimeError(f"Table with header '{header_text}' not found")
        
        # Select the range
        controller.select(table_range)
        
        # Force recalculation
        doc.calculateAll()
        
        # Export as PNG with selection only
        output_url = uno.systemPathToFileUrl(output_path)
        
        # Configure export - high quality PNG of selection only
        filter_data = PropertyValue(Name="PixelWidth", Value=1200)
        filter_data2 = PropertyValue(Name="PixelHeight", Value=900)
        
        export_props = (
            PropertyValue(Name="FilterName", Value="calc_png_Export"),
            PropertyValue(Name="SelectionOnly", Value=True),
        )
        
        doc.storeToURL(output_url, export_props)
        
    finally:
        doc.close(True)


def main():
    """Main entry point - read JSON from stdin, output result to stdout."""
    try:
        input_data = json.load(sys.stdin)
        
        excel_base64 = input_data['excelBase64']
        table_name = input_data['tableName']
        
        # Map table names to likely sheet names
        sheet_mapping = {
            'Sources and Uses': 'S&U ',
            'Take Out Loan Sizing': 'S&U ',
            'Capital Stack at Closing': 'S&U ',
            'Loan to Cost': 'LTC and LTV Calcs',
            'Loan to Value': 'LTC and LTV Calcs',
            'PILOT Schedule': 'S&U ',
        }
        
        sheet_name = sheet_mapping.get(table_name, 'S&U ')
        
        # Save Excel to temp file
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
            f.write(base64.b64decode(excel_base64))
            excel_path = f.name
        
        output_path = tempfile.mktemp(suffix='.png')
        
        try:
            # Connect to LibreOffice
            desktop = connect_to_libreoffice()
            
            # Export the table
            export_range_as_image(desktop, excel_path, sheet_name, table_name, output_path)
            
            # Read result and return
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