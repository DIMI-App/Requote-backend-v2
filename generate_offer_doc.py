import json
import os
from docx import Document

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OFFER_2_PATH = os.path.join(BASE_DIR, "offer2_template.docx")
OFFER_1_DATA_PATH = os.path.join(BASE_DIR, "outputs", "items_offer1.json")
OUTPUT_PATH = os.path.join(BASE_DIR, "outputs", "final_offer1.docx")

print("Starting offer generation")

try:
    with open(OFFER_1_DATA_PATH, "r", encoding="utf-8") as f:
        full_data = json.load(f)
    
    items = full_data.get("items", [])
    print(f"Loaded {len(items)} items")
    
    if len(items) == 0:
        print("No items found")
        exit(1)

except Exception as e:
    print(f"Error loading data: {e}")
    exit(1)

try:
    doc = Document(OFFER_2_PATH)
    print(f"Template loaded with {len(doc.tables)} tables")
except Exception as e:
    print(f"Error loading template: {e}")
    exit(1)

if len(doc.tables) == 0:
    print("No tables in template")
    exit(1)

best_table = None
best_score = 0

for idx, table in enumerate(doc.tables):
    if len(table.rows) < 2:
        continue
    
    header_text = ' '.join([cell.text.lower() for cell in table.rows[0].cells])
    
    score = 0
    if 'description' in header_text or 'item' in header_text:
        score += 1
    if 'price' in header_text:
        score += 1
    if 'quantity' in header_text:
        score += 1
    
    if score > best_score:
        best_score = score
        best_table = table

if best_table is None:
    best_table = doc.tables[0]

print(f"Using table with {len(best_table.rows)} rows")

header_row = best_table.rows[0]
num_cols = len(best_table.columns)

col_map = {}
for col_idx, cell in enumerate(header_row.cells):
    text = cell.text.lower()
    if 'description' in text or 'item' in text or 'name' in text:
        col_map['description'] = col_idx
    elif 'price' in text:
        col_map['price'] = col_idx
    elif 'quantity' in text or 'qty' in text:
        col_map['quantity'] = col_idx

if 'description' not in col_map:
    col_map['description'] = 0

while len(best_table.rows) > 1:
    best_table._tbl.remove(best_table.rows[1]._tr)

print(f"Inserting {len(items)} items")

for idx, item in enumerate(items, start=1):
    row = best_table.add_row().cells
    
    try:
        desc = item.get("item_name", "")
        if item.get("details"):
            desc = desc + "\n" + item.get("details")
        
        if 'description' in col_map and col_map['description'] < len(row):
            row[col_map['description']].text = desc
        
        if 'price' in col_map and col_map['price'] < len(row):
            row[col_map['price']].text = str(item.get("unit_price", ""))
        
        if 'quantity' in col_map and col_map['quantity'] < len(row):
            row[col_map['quantity']].text = str(item.get("quantity", "1"))
    
    except Exception as e:
        print(f"Error on item {idx}: {e}")
        continue

try:
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    doc.save(OUTPUT_PATH)
    print(f"Success! Saved to {OUTPUT_PATH}")
except Exception as e:
    print(f"Error saving: {e}")
    exit(1)

print("Done")
```

---

## ✅ What I did:

1. **Removed ALL special characters** - No emojis, no fancy unicode
2. **Removed ALL f-strings with special chars** - Plain strings only
3. **Simplified everything** - Minimal code, maximum compatibility
4. **No fancy formatting** - Just basic print statements

---

## 📝 Do this:

1. **Copy the code above**
2. **Open VS Code**
3. **Open `generate_offer_doc.py`**
4. **DELETE EVERYTHING**
5. **Paste the new code**
6. **Save** (Ctrl+S)
7. **Push to GitHub:**
```
   Day 13: Ultra-clean generate_offer_doc.py