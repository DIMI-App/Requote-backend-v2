import json
import os
from docx import Document

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OFFER_2_PATH = os.path.join(BASE_DIR, "offer2_template.docx")
OFFER_1_DATA_PATH = os.path.join(BASE_DIR, "outputs", "items_offer1.json")
OUTPUT_PATH = os.path.join(BASE_DIR, "outputs", "final_offer1.docx")

print("=" * 60)
print("REQUOTE AI - OFFER GENERATOR")
print("=" * 60)

# Load JSON data
print("\nStep 1: Loading extracted data...")
try:
    with open(OFFER_1_DATA_PATH, "r", encoding="utf-8") as f:
        full_data = json.load(f)
    
    items = full_data.get("items", [])
    
    print(f"Loaded {len(items)} items")
    
    if len(items) == 0:
        print("ERROR: No items found")
        exit(1)
    
except Exception as e:
    print(f"ERROR loading data: {e}")
    exit(1)

# Load template
print("\nStep 2: Loading template...")
try:
    doc = Document(OFFER_2_PATH)
    print(f"Template loaded ({len(doc.tables)} tables)")
except Exception as e:
    print(f"ERROR loading template: {e}")
    exit(1)

if len(doc.tables) == 0:
    print("ERROR: No tables in template")
    exit(1)

# Find products table
print("\nStep 3: Finding products table...")

def find_products_table(doc):
    keywords = ['position', 'description', 'price', 'quantity', 'total', 'item']
    
    best_match = None
    best_score = 0
    
    for idx, table in enumerate(doc.tables):
        if len(table.rows) < 2:
            continue
        
        header_text = ' '.join([cell.text.lower() for cell in table.rows[0].cells])
        score = sum(1 for kw in keywords if kw in header_text)
        
        if len(table.columns) >= 3:
            score += 2
        
        if score > best_score:
            best_score = score
            best_match = (idx, table)
    
    if best_match is None:
        largest = max(range(len(doc.tables)), 
                     key=lambda i: len(doc.tables[i].rows) * len(doc.tables[i].columns))
        best_match = (largest, doc.tables[largest])
    
    return best_match

table_idx, product_table = find_products_table(doc)
print(f"Using table {table_idx} ({len(product_table.rows)}x{len(product_table.columns)})")

# Detect columns
print("\nStep 4: Detecting columns...")

def detect_columns(table):
    header_row = table.rows[0]
    mapping = {}
    
    for col_idx, cell in enumerate(header_row.cells):
        text = cell.text.lower().strip()
        
        if 'position' in text or 'num' in text:
            mapping['position'] = col_idx
        elif 'description' in text or 'name' in text or 'item' in text:
            mapping['description'] = col_idx
        elif 'quantity' in text or 'qty' in text:
            mapping['quantity'] = col_idx
        elif 'price' in text and 'unit' in text:
            mapping['unit_price'] = col_idx
        elif 'total' in text or 'sum' in text:
            mapping['total'] = col_idx
    
    num_cols = len(table.columns)
    
    if 'description' not in mapping:
        mapping['description'] = 1 if num_cols > 1 else 0
    if 'position' not in mapping and num_cols > 1:
        mapping['position'] = 0
    if 'quantity' not in mapping and num_cols > 2:
        mapping['quantity'] = 2
    if 'unit_price' not in mapping and num_cols > 3:
        mapping['unit_price'] = 3
    if 'total' not in mapping and num_cols > 4:
        mapping['total'] = 4
    
    return mapping

column_map = detect_columns(product_table)
print(f"Columns: {column_map}")

# Clear old rows
print("\nStep 5: Clearing old rows...")
original_count = len(product_table.rows)

while len(product_table.rows) > 1:
    product_table._tbl.remove(product_table.rows[1]._tr)

print(f"Cleared {original_count - 1} rows")

# Insert items
print(f"\nStep 6: Inserting {len(items)} items...")

for idx, item in enumerate(items, start=1):
    row = product_table.add_row().cells
    
    try:
        if 'position' in column_map:
            row[column_map['position']].text = str(idx)
        
        if 'description' in column_map:
            desc_parts = []
            if item.get("item_name"):
                desc_parts.append(item["item_name"])
            if item.get("details"):
                desc_parts.append(item["details"])
            
            description = "\n".join(desc_parts) if desc_parts else ""
            row[column_map['description']].text = description
        
        if 'quantity' in column_map and item.get("quantity"):
            row[column_map['quantity']].text = str(item["quantity"])
        
        if 'unit_price' in column_map and item.get("unit_price"):
            row[column_map['unit_price']].text = str(item["unit_price"])
        
        if 'total' in column_map and item.get("total_price"):
            row[column_map['total']].text = str(item["total_price"])
    
    except Exception as e:
        print(f"Warning: Error on item {idx}: {e}")
        continue

print(f"Inserted {len(items)} items")

# Save
print("\nStep 7: Saving...")
try:
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    doc.save(OUTPUT_PATH)
    print(f"SUCCESS! Saved to: {OUTPUT_PATH}")
except Exception as e:
    print(f"ERROR saving: {e}")
    exit(1)

print("\n" + "=" * 60)
print("COMPLETED")
print("=" * 60)
```

---

## 📝 Do this:

1. **Replace `generate_offer_doc.py`** with code above
2. **Save**
3. **Push:**
```
   Day 13: Improve error handling in generate_offer_doc.py