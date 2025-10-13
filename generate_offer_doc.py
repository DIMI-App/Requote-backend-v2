import json
import os
from docx import Document

# === FILE PATHS ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OFFER_2_PATH = os.path.join(BASE_DIR, "offer2_template.docx")
OFFER_1_DATA_PATH = os.path.join(BASE_DIR, "outputs", "items_offer1.json")
OUTPUT_PATH = os.path.join(BASE_DIR, "outputs", "final_offer1.docx")

print("=" * 60)
print("🚀 REQUOTE AI - INTELLIGENT HYBRID GENERATOR (SV3)")
print("=" * 60)

# === STEP 1: Load complete data from JSON ===
print("\n📂 Step 1: Loading extracted data...")
try:
    with open(OFFER_1_DATA_PATH, "r", encoding="utf-8") as f:
        full_data = json.load(f)
    
    # Extract components
    items = full_data.get("items", [])
    technical_specs = full_data.get("technical_specs", {})
    company_info = full_data.get("company_info", {})
    
    print(f"✅ Loaded data:")
    print(f"   • Items: {len(items)}")
    
    # === Safety check for empty items ===
    if len(items) == 0:
        print("\n❌ CRITICAL ERROR: NO ITEMS FOUND")
        print("   Please check extracted data")
        exit(1)
    
except Exception as e:
    print(f"❌ ERROR loading data: {e}")
    exit(1)

# === STEP 2: Load DOCX Template ===
print("\n📄 Step 2: Loading template...")
try:
    doc = Document(OFFER_2_PATH)
    print(f"✅ Template loaded ({len(doc.tables)} tables found)")
except Exception as e:
    print(f"❌ ERROR loading template: {e}")
    exit(1)

if len(doc.tables) == 0:
    print("❌ ERROR: No tables found in template")
    exit(1)

# === STEP 3: Find products table ===
print("\n🔍 Step 3: Detecting products table...")

def find_products_table(doc):
    best_match = None
    best_score = 0
    
    product_keywords = [
        'позиція', 'опис', 'ціна', 'кількість', 'сума',
        'position', 'description', 'price', 'quantity', 'total',
        'item', 'product'
    ]
    
    for table_idx, table in enumerate(doc.tables):
        if len(table.rows) < 2:
            continue
        
        header_row = table.rows[0]
        header_text = ' '.join([cell.text.lower() for cell in header_row.cells])
        
        score = sum(1 for keyword in product_keywords if keyword in header_text)
        
        if len(table.columns) >= 3:
            score += 2
        
        print(f"   Table {table_idx}: {len(table.rows)}x{len(table.columns)} - Score: {score}")
        
        if score > best_score:
            best_score = score
            best_match = (table_idx, table)
    
    if best_match is None:
        largest_idx = max(range(len(doc.tables)), 
                         key=lambda i: len(doc.tables[i].rows) * len(doc.tables[i].columns))
        best_match = (largest_idx, doc.tables[largest_idx])
    
    return best_match

table_idx, product_table = find_products_table(doc)
print(f"✅ Selected Table #{table_idx} ({len(product_table.rows)}x{len(product_table.columns)})")

# === STEP 4: Column mapping ===
print("\n🗂️  Step 4: Detecting column structure...")

def detect_columns(table):
    header_row = table.rows[0]
    mapping = {}
    
    for col_idx, cell in enumerate(header_row.cells):
        text = cell.text.lower().strip()
        
        if not mapping.get('position') and any(k in text for k in ['№', 'поз', 'position', 'num']):
            mapping['position'] = col_idx
        elif not mapping.get('description') and any(k in text for k in ['опис', 'description', 'name', 'item']):
            mapping['description'] = col_idx
        elif not mapping.get('quantity') and any(k in text for k in ['кількість', 'quantity', 'qty']):
            mapping['quantity'] = col_idx
        elif not mapping.get('unit_price') and any(k in text for k in ['ціна', 'price', 'unit']):
            mapping['unit_price'] = col_idx
        elif not mapping.get('total') and any(k in text for k in ['сума', 'total', 'sum']):
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

print("📋 Column mapping:")
for field, col in sorted(column_map.items()):
    print(f"   • {field}: column {col}")

# === STEP 5: Clear old rows ===
print(f"\n🧹 Step 5: Clearing old rows...")
original_count = len(product_table.rows)

while len(product_table.rows) > 1:
    product_table._tbl.remove(product_table.rows[1]._tr)

print(f"✅ Cleared {original_count - 1} rows, kept header")

# === STEP 6: Insert items ===
print(f"\n📝 Step 6: Inserting {len(items)} items...")

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
    
    except IndexError as e:
        print(f"   ⚠️  Warning: Column index out of range for item {idx}")
        continue

print(f"✅ Successfully inserted {len(items)} items")

# === STEP 7: Save ===
print("\n💾 Step 7: Saving final offer...")
try:
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    doc.save(OUTPUT_PATH)
    print(f"✅ SUCCESS! Saved to: {OUTPUT_PATH}")
except Exception as e:
    print(f"❌ ERROR saving document: {e}")
    exit(1)

print("\n" + "=" * 60)
print("📊 SUMMARY:")
print(f"   • Items processed: {len(items)}")
print(f"   • Table used: #{table_idx}")
print(f"   • Columns mapped: {len(column_map)}")
print("=" * 60)
print("✨ Done! Your branded offer is ready (SV3)")
print("=" * 60)