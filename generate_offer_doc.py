import json
import os
from docx import Document

# === FILE PATHS ===
OFFER_2_PATH = "offer2_template.docx"
OFFER_1_DATA_PATH = "outputs/items_offer1.json"
OUTPUT_PATH = "outputs/final_offer1.docx"

print("🚀 Starting Requote AI - Offer Generator\n")

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

# === STEP 3: Find the products table ===
print("🔍 Looking for products table...")

# We know TABLE #2 (index 1) is the products table
if len(doc.tables) < 2:
    raise Exception("❌ Not enough tables in document!")

product_table = doc.tables[1]  # Second table (index 1)
print(f"✅ Found products table with {len(product_table.rows)} rows\n")

# === STEP 4: Clear old product rows (keep header) ===
print("🧹 Clearing old product data...")
original_rows = len(product_table.rows)

# Keep first row (header), delete everything else
while len(product_table.rows) > 1:
    product_table._tbl.remove(product_table.rows[1]._tr)

print(f"✅ Cleared {original_rows - 1} old rows, header preserved\n")

# === STEP 5: Add new rows from Offer 1 ===
print(f"📝 Adding {len(items)} new items from Offer 1...")

for idx, item in enumerate(items, start=1):
    row = product_table.add_row().cells
    
    # Column 0: Position number
    row[0].text = str(idx)
    
    # Column 1: Description (combine name + description)
    name = item.get("name", "")
    description = item.get("description", "")
    
    if name and description:
        full_description = f"{name}\n{description}"
    elif name:
        full_description = name
    elif description:
        full_description = description
    else:
        full_description = ""
    
    row[1].text = str(full_description)
    
    # Column 2: Price (convert to string safely)
    price = item.get("price", "")
    row[2].text = str(price) if price else ""
    
    # Column 3: Quantity (convert to string safely)
    quantity = item.get("quantity", "")
    row[3].text = str(quantity) if quantity else ""
    
    # Column 4: Total (leave empty for now)
    row[4].text = ""

print(f"✅ Successfully added {len(items)} items\n")

# === STEP 6: Save new document ===
print("💾 Saving final offer...")
os.makedirs("outputs", exist_ok=True)
doc.save(OUTPUT_PATH)

print(f"✅ SUCCESS! Final offer saved to: {OUTPUT_PATH}")
print(f"\n📊 Summary:")
print(f"   • Items processed: {len(items)}")
print(f"   • Output file: {OUTPUT_PATH}")
print(f"   • Table rows: {len(product_table.rows)} (1 header + {len(items)} items)")
print("\n✨ Done! Open the file to check your new offer.")