import os
import json
import openpyxl
from openpyxl.styles import Font, Alignment

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OFFER_2_PATH = os.path.join(BASE_DIR, "offer2_template.xlsx")
ITEMS_PATH = os.path.join(BASE_DIR, "outputs", "items_offer1.json")
OUTPUT_PATH = os.path.join(BASE_DIR, "outputs", "final_offer1.xlsx")

print("=" * 60)
print("GENERATE OFFER FROM XLSX TEMPLATE")
print("=" * 60)

# Load items
try:
    with open(ITEMS_PATH, "r", encoding="utf-8") as f:
        full_data = json.load(f)
    items = full_data.get("items", [])
    print(f"✓ Loaded {len(items)} items")
except Exception as e:
    print(f"✗ Error loading items: {str(e)}")
    exit(1)

# Load Excel template
try:
    workbook = openpyxl.load_workbook(OFFER_2_PATH)
    sheet = workbook.active
    print(f"✓ Template loaded: {sheet.title}")
except Exception as e:
    print(f"✗ Error loading template: {str(e)}")
    exit(1)

# Find the pricing table header row
table_start_row = None
for idx, row in enumerate(sheet.iter_rows(min_row=1, max_row=50), start=1):
    row_text = ' '.join([str(cell.value) if cell.value else '' for cell in row]).upper()
    if any(keyword in row_text for keyword in ['POSITION', 'DESCRIPTION', 'PRICE', 'QUANTITY', 'TOTAL']):
        table_start_row = idx
        print(f"✓ Found pricing table header at row {table_start_row}")
        break

if not table_start_row:
    print("✗ Could not find pricing table header")
    exit(1)

# Clear existing data rows (keep header)
data_start_row = table_start_row + 1
max_row = sheet.max_row

print(f"Clearing rows {data_start_row} to {max_row}...")
for row_idx in range(max_row, data_start_row - 1, -1):
    sheet.delete_rows(row_idx)

print(f"✓ Cleared existing data rows")

# Insert items into Excel
current_row = data_start_row
item_number = 1

for item in items:
    category = item.get('category', '')
    
    # Add category row if present
    if category and (current_row == data_start_row or category != last_category):
        sheet.insert_rows(current_row)
        # Merge cells for category (across all columns)
        first_col = sheet.cell(table_start_row, 1).column
        last_col = sheet.cell(table_start_row, sheet.max_column).column
        
        category_cell = sheet.cell(current_row, 2)  # Usually description column
        category_cell.value = category
        category_cell.font = Font(bold=True, size=11)
        category_cell.alignment = Alignment(horizontal='left')
        
        current_row += 1
        last_category = category
    
    # Add item row
    sheet.insert_rows(current_row)
    
    # Column 1: Item number
    sheet.cell(current_row, 1).value = str(item_number)
    item_number += 1
    
    # Column 2: Description
    desc = item.get('item_name', '')
    if item.get('details'):
        desc += '\n' + item.get('details')
    sheet.cell(current_row, 2).value = desc
    sheet.cell(current_row, 2).alignment = Alignment(wrap_text=True, vertical='top')
    
    # Column 3: Unit Price
    unit_price = item.get('unit_price', '')
    sheet.cell(current_row, 3).value = unit_price
    sheet.cell(current_row, 3).alignment = Alignment(horizontal='right')
    
    # Column 4: Quantity
    qty = item.get('quantity', '1')
    sheet.cell(current_row, 4).value = qty
    sheet.cell(current_row, 4).alignment = Alignment(horizontal='center')
    
    # Column 5: Total Price
    total_price = item.get('total_price', unit_price)
    sheet.cell(current_row, 5).value = total_price
    sheet.cell(current_row, 5).alignment = Alignment(horizontal='right')
    
    current_row += 1

last_category = None

print(f"✓ Inserted {len(items)} items into Excel")

# Save output
try:
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    workbook.save(OUTPUT_PATH)
    file_size = os.path.getsize(OUTPUT_PATH)
    print(f"✓ Saved: {OUTPUT_PATH}")
    print(f"  File size: {file_size:,} bytes")
except Exception as e:
    print(f"✗ Error saving: {str(e)}")
    exit(1)

print("=" * 60)
print("GENERATION COMPLETE")
print("=" * 60)