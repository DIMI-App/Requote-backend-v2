import json
import os
from docx import Document
import re

# === FILE PATHS ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OFFER_2_PATH = os.path.join(BASE_DIR, "offer2_template.docx")
OFFER_1_DATA_PATH = os.path.join(BASE_DIR, "outputs", "items_offer1.json")
OUTPUT_PATH = os.path.join(BASE_DIR, "outputs", "final_offer1.docx")

print("🚀 Starting Requote AI - Intelligent Offer Generator\n")

# === STEP 1: Load items from JSON ===
print("📂 Loading extracted items from Offer 1...")
with open(OFFER_1_DATA_PATH, "r", encoding="utf-8") as f:
    try:
        data = f.read().strip()
        items = json.loads(data)
        print(f"✅ Loaded {len(items)} items\n")
    except json.JSONDecodeError as e:
        print("❌ JSON loading failed:", e)
        raise

# === STEP 2: Load DOCX Template (Offer 2) ===
print("📄 Loading Offer 2 template...")
doc = Document(OFFER_2_PATH)
print(f"✅ Template loaded ({len(doc.tables)} tables found)\n")

# === STEP 3: Find the products table intelligently ===
print("🔍 Scanning for products table...")

def detect_product_table(doc):
    """
    Scan all tables and find the one that looks like a products/items table
    """
    for table_index, table in enumerate(doc.tables):
        if len(table.rows) < 2:
            continue  # Skip tables with less than 2 rows
        
        # Get first row (potential header)
        header_row = table.rows[0]
        header_text = ' '.join([cell.text.lower() for cell in header_row.cells])
        
        # Check if this looks like a products table
        product_keywords = ['позиція', 'опис', 'ціна', 'кількість', 'сума', 
                           'position', 'description', 'price', 'quantity', 'total',
                           'item', 'product', 'наименование', 'товар']
        
        matches = sum(1 for keyword in product_keywords if keyword in header_text)
        
        if matches >= 2:  # At least 2 keywords match
            print(f"✅ Found products table at index {table_index}")
            print(f"   Header: {header_text[:100]}...")
            return table_index, table
    
    # If no clear match, use the largest table (most likely to be products)
    largest_table_index = max(range(len(doc.tables)), 
                              key=lambda i: len(doc.tables[i].rows) * len(doc.tables[i].columns))
    print(f"⚠️  No clear products table found, using largest table at index {largest_table_index}")
    return largest_table_index, doc.tables[largest_table_index]

table_index, product_table = detect_product_table(doc)
print(f"📊 Table has {len(product_table.rows)} rows × {len(product_table.columns)} columns\n")

# === STEP 4: Detect column mapping ===
print("🗂️  Detecting column structure...")

def detect_column_mapping(table):
    """
    Analyze header row to determine which column is for what
    Returns dict like: {'description': 1, 'quantity': 3, 'price': 2}
    """
    header_row = table.rows[0]
    mapping = {}
    
    for col_index, cell in enumerate(header_row.cells):
        cell_text = cell.text.lower().strip()
        
        # Position/Number column
        if any(word in cell_text for word in ['поз', 'position', '№', 'num', 'п/п']):
            mapping['position'] = col_index
        
        # Description column
        elif any(word in cell_text for word in ['опис', 'description', 'найменування', 'наименование', 'товар', 'item', 'product']):
            mapping['description'] = col_index
        
        # Quantity column
        elif any(word in cell_text for word in ['кількість', 'quantity', 'кіл-сть', 'qty', 'к-сть']):
            mapping['quantity'] = col_index
        
        # Price column
        elif any(word in cell_text for word in ['ціна', 'price', 'вартість', 'unit']):
            mapping['price'] = col_index
        
        # Total column
        elif any(word in cell_text for word in ['сума', 'total', 'разом', 'всього']):
            mapping['total'] = col_index
    
    print("📋 Column mapping detected:")
    for field, col in mapping.items():
        print(f"   {field}: column {col}")
    
    return mapping

column_map = detect_column_mapping(product_table)

# If key columns not found, use defaults
if 'description' not in column_map and len(product_table.columns) > 1:
    column_map['description'] = 1  # Usually second column
    print("⚠️  Using default: description = column 1")

print()

# === STEP 5: Clear old rows (keep header) ===
print("🧹 Clearing old product data...")
original_rows = len(product_table.rows)

# Keep first row (header), delete everything else
while len(product_table.rows) > 1:
    product_table._tbl.remove(product_table.rows[1]._tr)

print(f"✅ Cleared {original_rows - 1} old rows, header preserved\n")

# === STEP 6: Add new rows from Offer 1 ===
print(f"📝 Adding {len(items)} new items from Offer 1...")

for idx, item in enumerate(items, start=1):
    row = product_table.add_row().cells
    
    # Position number
    if 'position' in column_map:
        row[column_map['position']].text = str(idx)
    
    # Description (combine name + details)
    if 'description' in column_map:
        item_name = item.get("item_name", "")
        details = item.get("details", "")
        
        if item_name and details:
            full_description = f"{item_name}\n{details}"
        elif item_name:
            full_description = item_name
        elif details:
            full_description = details
        else:
            full_description = ""
        
        row[column_map['description']].text = full_description
    
    # Quantity
    if 'quantity' in column_map:
        quantity = item.get("quantity", "")
        row[column_map['quantity']].text = str(quantity)
    
    # Unit Price
    if 'price' in column_map:
        unit_price = item.get("unit_price", "")
        row[column_map['price']].text = str(unit_price)
    
    # Total Price
    if 'total' in column_map:
        total_price = item.get("total_price", "")
        row[column_map['total']].text = str(total_price)

print(f"✅ Successfully added {len(items)} items\n")

# === STEP 7: Save new document ===
print("💾 Saving final offer...")
os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
doc.save(OUTPUT_PATH)

print(f"✅ SUCCESS! Final offer saved to: {OUTPUT_PATH}")
print(f"\n📊 Summary:")
print(f"   • Items processed: {len(items)}")
print(f"   • Output file: {OUTPUT_PATH}")
print(f"   • Table rows: {len(product_table.rows)} (1 header + {len(items)} items)")
print("\n✨ Done! Your branded offer is ready.")