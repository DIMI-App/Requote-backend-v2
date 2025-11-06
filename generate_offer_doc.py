import os
import json
from docx import Document
from docx.shared import Pt

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OFFER_2_PATH = os.path.join(BASE_DIR, "offer2_template.docx")
ITEMS_PATH = os.path.join(BASE_DIR, "outputs", "items_offer1.json")
OUTPUT_PATH = os.path.join(BASE_DIR, "outputs", "final_offer1.docx")

print("=" * 60, flush=True)
print("GENERATE OFFER - Starting", flush=True)
print("=" * 60, flush=True)

# Load items
try:
    print(f"Loading items from: {ITEMS_PATH}", flush=True)
    with open(ITEMS_PATH, "r", encoding="utf-8") as f:
        full_data = json.load(f)
    
    items = full_data.get("items", [])
    print(f"✓ Loaded {len(items)} items", flush=True)
    
    if len(items) == 0:
        print("✗ No items found", flush=True)
        exit(1)

except Exception as e:
    print(f"✗ Error loading items: {str(e)}", flush=True)
    exit(1)

# Load template
try:
    print(f"Loading template from: {OFFER_2_PATH}", flush=True)
    doc = Document(OFFER_2_PATH)
    print(f"✓ Template loaded: {len(doc.tables)} tables", flush=True)
except Exception as e:
    print(f"✗ Error loading template: {str(e)}", flush=True)
    exit(1)

if len(doc.tables) == 0:
    print("✗ No tables in template", flush=True)
    exit(1)

# Find pricing table (usually table with most columns)
best_table = None
max_cols = 0

for idx, table in enumerate(doc.tables):
    cols = len(table.columns)
    print(f"Table {idx + 1}: {len(table.rows)} rows, {cols} columns", flush=True)
    if cols > max_cols and len(table.rows) > 1:
        max_cols = cols
        best_table = table

if best_table is None:
    print("✗ Could not find pricing table", flush=True)
    exit(1)

print(f"✓ Selected table with {len(best_table.columns)} columns", flush=True)

# Clear existing data rows (keep header)
print(f"Clearing {len(best_table.rows) - 1} existing rows...", flush=True)
while len(best_table.rows) > 1:
    best_table._tbl.remove(best_table.rows[1]._tr)

# Insert items
print(f"Inserting {len(items)} items...", flush=True)
for idx, item in enumerate(items, start=1):
    row = best_table.add_row().cells
    
    try:
        # Assume columns: Position, Description, Price, Quantity, Total
        if len(row) >= 1:
            row[0].text = f"{idx}."
        
        if len(row) >= 2:
            desc = item.get("item_name", "")
            if item.get("details"):
                desc = f"{desc}\n{item.get('details')}"
            row[1].text = desc
        
        if len(row) >= 3:
            price = str(item.get("unit_price", "")).replace("€", "").replace("$", "").strip()
            row[2].text = price
        
        if len(row) >= 4:
            row[3].text = str(item.get("quantity", "1"))
        
        if len(row) >= 5:
            total = str(item.get("total_price", "")).replace("€", "").replace("$", "").strip()
            row[4].text = total
        
        # Apply formatting
        for cell in row:
            if len(cell.paragraphs) > 0:
                para = cell.paragraphs[0]
                if len(para.runs) > 0:
                    para.runs[0].font.size = Pt(10)
        
        if idx % 10 == 0:
            print(f"  Inserted {idx} items...", flush=True)
    
    except Exception as e:
        print(f"  ✗ Error on item {idx}: {str(e)}", flush=True)
        continue

print(f"✓ Inserted all {len(items)} items", flush=True)

# Save document
try:
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    doc.save(OUTPUT_PATH)
    file_size = os.path.getsize(OUTPUT_PATH)
    print(f"✓ Document saved: {OUTPUT_PATH}", flush=True)
    print(f"  File size: {file_size:,} bytes", flush=True)
except Exception as e:
    print(f"✗ Error saving: {str(e)}", flush=True)
    exit(1)

print("=" * 60, flush=True)
print("GENERATION COMPLETE", flush=True)
print("=" * 60, flush=True)