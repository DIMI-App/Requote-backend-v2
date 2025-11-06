import os
import json
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OFFER_2_PATH = os.path.join(BASE_DIR, "offer2_template.docx")
ITEMS_PATH = os.path.join(BASE_DIR, "outputs", "items_offer1.json")
OUTPUT_PATH = os.path.join(BASE_DIR, "outputs", "final_offer1.docx")

def set_cell_background(cell, color):
    """Set cell background color"""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), color)
    cell._element.get_or_add_tcPr().append(shading_elm)

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

# Analyze header to understand column structure
header_row = best_table.rows[0]
headers = [cell.text.strip().lower() for cell in header_row.cells]
print(f"✓ Column headers: {headers}", flush=True)

# Clear existing data rows (keep header)
print(f"Clearing {len(best_table.rows) - 1} existing rows...", flush=True)
while len(best_table.rows) > 1:
    best_table._tbl.remove(best_table.rows[1]._tr)

# Group items by category
from collections import OrderedDict
categorized_items = OrderedDict()
for item in items:
    cat = item.get("category", "Main Items")
    if cat not in categorized_items:
        categorized_items[cat] = []
    categorized_items[cat].append(item)

print(f"✓ Grouped into {len(categorized_items)} categories", flush=True)

# Insert items with category separators
item_counter = 1
for category, cat_items in categorized_items.items():
    print(f"  Processing category: {category} ({len(cat_items)} items)", flush=True)
    
    # Add category header row (spanning all columns)
    category_row = best_table.add_row().cells
    merged_cell = category_row[0].merge(category_row[-1])
    merged_cell.text = category
    
    # Format category header
    for paragraph in merged_cell.paragraphs:
        paragraph.alignment = 1  # Center
        for run in paragraph.runs:
            run.bold = True
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(255, 255, 255)  # White text
    
    # Background color (blue)
    set_cell_background(merged_cell, "4472C4")
    
    # Add items in this category
    for item in cat_items:
        row = best_table.add_row().cells
        
        try:
            # Column 0: Position number
            if len(row) >= 1:
                row[0].text = f"{item_counter}."
                item_counter += 1
            
            # Column 1: Description
            if len(row) >= 2:
                desc = item.get("item_name", "")
                if item.get("details"):
                    desc = f"{desc}\n{item.get('details')}"
                row[1].text = desc
            
            # Column 2: Unit Price
            if len(row) >= 3:
                price = item.get("unit_price", "")
                # Preserve "Included" text or format numbers
                if "included" in str(price).lower():
                    row[2].text = "Included"
                else:
                    # Keep currency symbols and thousand separators
                    row[2].text = str(price)
            
            # Column 3: Quantity
            if len(row) >= 4:
                row[3].text = str(item.get("quantity", "1"))
            
            # Column 4: Total Price
            if len(row) >= 5:
                total = item.get("total_price", "")
                # Preserve "Included" text or format numbers
                if "included" in str(total).lower():
                    row[4].text = "Included"
                else:
                    row[4].text = str(total)
            
            # Apply formatting
            for cell in row:
                if len(cell.paragraphs) > 0:
                    para = cell.paragraphs[0]
                    if len(para.runs) > 0:
                        para.runs[0].font.size = Pt(10)
            
        except Exception as e:
            print(f"  ✗ Error on item: {str(e)}", flush=True)
            continue

print(f"✓ Inserted all items with {len(categorized_items)} category separators", flush=True)

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