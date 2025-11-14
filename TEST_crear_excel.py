#!/usr/bin/env python3
# TEST SCRIPT - Verify we can create Excel files in your directory

from openpyxl import Workbook
from datetime import datetime

print("="*60)
print("TEST: Creating Excel file in your directory")
print("="*60)

try:
    # Create simple workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "TEST"

    ws['A1'] = "TEST SUCCESSFUL"
    ws['A2'] = f"Created: {datetime.now()}"
    ws['A3'] = "Directory: YOUR LOCAL PATH"

    # Save file
    filename = "TEST_SUCCESS.xlsx"
    wb.save(filename)

    print(f"\n✅ SUCCESS! File created: {filename}")
    print(f"📁 Location: Current directory")
    print("\nIf you see this file, we can proceed with the full system.")
    print("="*60)

except Exception as e:
    print(f"\n❌ ERROR: {e}")
    print("Please install openpyxl: pip install openpyxl")

input("\nPress ENTER to close...")
