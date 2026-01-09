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


def find_table_in_all_sheets(doc, header_text):
    """
    Search ALL sheets for the header text.
    Returns (sheet, cell_range) or (None, None) if not found.
    """
    sheets = doc.getSheets()
    
    for i in range(sheets.getCount()):
        sheet = sheets.getByIndex(i)
        
        # Create search descriptor
        search = sheet.createSearchDescriptor()
        search.SearchString = header_text
        search.SearchCaseSensitive = False
        
        found = sheet.findFirst(search)
        if found:
            # Found the header - now expand to get full table
            table_range = expand_to_table(sheet, found)
            if table_range:
                return sheet, table_range
    
    return None, None


def expand_to_table(sheet, header_cell):
    """
    Starting from header cell, expand to find full table boundaries.
    """
    start_col = header_cell.CellAddress.Column
    start_row = header_cell.CellAddress.Row
    
    # Find table width by scanning header row
    end_col = start_col
    col = start_col
    while col < start_col + 50:  # Max 50 columns
        cell = sheet.getCellByPosition(col, start_row)
        cell_type = cell.getType()  # 0=EMPTY, 1=VALUE, 2=STRING, 3=FORMULA
        cell_str = cell.getString().strip()
        
        if col > start_col and cell_type == 0 and not cell_str:
            # Empty cell - check if next 2 are also empty (end of table)
            next_empty = True
            for check in range(1, 3):
                if col + check < start_col + 50:
                    next_cell = sheet.getCellByPosition(col + check, start_row)
                    if next_cell.getType() != 0 or next_cell.getString().strip():
                        next_empty = False
                        break
            if next_empty:
                break
        
        if cell_type != 0 or cell_str:
            end_col = col
        col += 1
    
    # Find table height by scanning down
    end_row = start_row
    row = start_row
    consecutive_empty = 0
    
    while row < start_row + 100 and consecutive_empty < 2:  # Max 100 rows
        row += 1
        
        # Check if entire row is empty
        row_has_content = False
        for c in range(start_col, end_col + 1):
            cell = sheet.getCellByPosition(c, row)
            if cell.getType() != 0 or cell.getString().strip():
                row_has_content = True
                break
        
        if row_has_content:
            end_row = row
            consecutive_empty = 0
        else:
            consecutive_empty += 1
    
    return sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)


def export_range_as_image(doc, sheet, table_range, output_path):
    """
    Select range and export as PNG.
    """
    # Activate the sheet
    controller = doc.getCurrentController()
    controller.setActiveSheet(sheet)
    
    # Select the range
    controller.select(table_range)
    
    # Force recalculation
    doc.calculateAll()
    
    # Export as PNG with selection only
    output_url = uno.systemPathToFileUrl(output_path)
    
    export_props = (
        PropertyValue(Name="FilterName", Value="calc_png_Export"),
        PropertyValue(Name="SelectionOnly", Value=True),
    )
    
    doc.storeToURL(output_url, export_props)


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
            
            # Open the document
            file_url = uno.systemPathToFileUrl(excel_path)
            load_props = (
                PropertyValue(Name="Hidden", Value=True),
                PropertyValue(Name="ReadOnly", Value=True),
            )
            doc = desktop.loadComponentFromURL(file_url, "_blank", 0, load_props)
            
            if not doc:
                raise RuntimeError("Failed to open document")
            
            try:
                # Search ALL sheets for the table header
                sheet, table_range = find_table_in_all_sheets(doc, table_name)
                
                if not sheet or not table_range:
                    # Table not found - return gracefully
                    print(json.dumps({
                        'success': False,
                        'error': f"Table '{table_name}' not found in any sheet"
                    }))
                    return
                
                # Export the range
                export_range_as_image(doc, sheet, table_range, output_path)
                
                # Read result
                with open(output_path, 'rb') as f:
                    image_base64 = base64.b64encode(f.read()).decode('utf-8')
                
                print(json.dumps({
                    'success': True,
                    'image': image_base64
                }))
                
            finally:
                doc.close(True)
                
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