import json
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OFFER_2_PATH = os.path.join(BASE_DIR, "offer2_template.docx")
OFFER_1_DATA_PATH = os.path.join(BASE_DIR, "outputs", "items_offer1.json")
OUTPUT_PATH = os.path.join(BASE_DIR, "outputs", "final_offer1.docx")

print("=" * 60)
print("SMART OFFER GENERATION - SV6")
print("=" * 60)

# Load extracted items
try:
    with open(OFFER_1_DATA_PATH, "r", encoding="utf-8") as f:
        full_data = json.load(f)
    
    items = full_data.get("items", [])
    print(f"✓ Loaded {len(items)} items from extraction")
    
    if len(items) == 0:
        print("✗ No items found in extraction")
        exit(1)

except Exception as e:
    print(f"✗ Error loading data: {str(e)}")
    exit(1)

# Load template
try:
    doc = Document(OFFER_2_PATH)
    print(f"✓ Template loaded: {len(doc.tables)} tables found")
except Exception as e:
    print(f"✗ Error loading template: {str(e)}")
    exit(1)

if len(doc.tables) == 0:
    print("✗ No tables in template")
    exit(1)

# SMART TABLE DETECTION
def analyze_table(table):
    """Analyze table to determine if it's a pricing table"""
    if len(table.rows) < 2:
        return 0, {}
    
    header_text = ' '.join([cell.text.lower() for cell in table.rows[0].cells])
    
    score = 0
    col_map = {}
    
    # Check for pricing indicators in ANY language
    pricing_keywords = ['price', 'ціна', 'precio', 'preis', 'prix', 'prezzo', 'cost', 'amount', 'сума']
    desc_keywords = ['description', 'опис', 'descripción', 'beschreibung', 'item', 'product', 'позиція']
    qty_keywords = ['quantity', 'кількість', 'cantidad', 'menge', 'quantité', 'qty', 'amt']
    
    for keyword in pricing_keywords:
        if keyword in header_text:
            score += 2
            break
    
    for keyword in desc_keywords:
        if keyword in header_text:
            score += 2
            break
    
    for keyword in qty_keywords:
        if keyword in header_text:
            score += 1
            break
    
    # Map columns by content, not just names
    for col_idx, cell in enumerate(table.rows[0].cells):
        text = cell.text.lower().strip()
        
        # Position column (usually first, short)
        if col_idx == 0 and (len(text) < 10 or 'поз' in text or 'no' in text or '#' in text):
            col_map['position'] = col_idx
        
        # Description column (look for keywords)
        if any(kw in text for kw in desc_keywords):
            col_map['description'] = col_idx
        
        # Price column
        if any(kw in text for kw in pricing_keywords) and 'сума' not in text:
            col_map['price'] = col_idx
        
        # Quantity column
        if any(kw in text for kw in qty_keywords):
            col_map['quantity'] = col_idx
        
        # Total/Sum column (usually last)
        if 'сума' in text or 'total' in text or 'sum' in text or 'amount' in text:
            col_map['total'] = col_idx
    
    # Smart fallback: assume standard table structure
    if 'description' not in col_map and len(table.columns) >= 2:
        col_map['description'] = 1  # Usually second column
    
    if 'price' not in col_map and len(table.columns) >= 3:
        col_map['price'] = 2  # Usually third column
    
    return score, col_map

# Find the best pricing table
best_table = None
best_score = 0
best_col_map = {}

print("\nAnalyzing tables...")
for idx, table in enumerate(doc.tables):
    score, col_map = analyze_table(table)
    print(f"  Table {idx + 1}: score={score}, columns={len(table.columns)}, mapped={len(col_map)} cols")
    
    if score > best_score:
        best_score = score
        best_table = table
        best_col_map = col_map

if best_table is None:
    print("✗ Could not find pricing table")
    exit(1)

print(f"\n✓ Selected pricing table with score {best_score}")
print(f"  Column mapping: {best_col_map}")

# Preserve header row formatting
header_row = best_table.rows[0]
header_cells_format = []

for cell in header_row.cells:
    cell_format = {
        'text': cell.text,
        'bold': False,
        'font_size': None,
        'alignment': None
    }
    
    if len(cell.paragraphs) > 0:
        para = cell.paragraphs[0]
        if len(para.runs) > 0:
            run = para.runs[0]
            cell_format['bold'] = run.bold
            if run.font.size:
                cell_format['font_size'] = run.font.size
        cell_format['alignment'] = para.alignment
    
    header_cells_format.append(cell_format)

# Clear existing data rows (keep header)
print(f"\nClearing {len(best_table.rows) - 1} existing data rows...")
while len(best_table.rows) > 1:
    best_table._tbl.remove(best_table.rows[1]._tr)

# Insert items
print(f"\nInserting {len(items)} items...")
for idx, item in enumerate(items, start=1):
    row = best_table.add_row().cells
    
    try:
        # Position column
        if 'position' in best_col_map and best_col_map['position'] < len(row):
            row[best_col_map['position']].text = f"{idx}."
        
        # Description column
        if 'description' in best_col_map and best_col_map['description'] < len(row):
            desc = item.get("item_name", "")
            if item.get("details"):
                desc = f"{desc}\n{item.get('details')}"
            row[best_col_map['description']].text = desc
        
        # Price column
        if 'price' in best_col_map and best_col_map['price'] < len(row):
            price = str(item.get("unit_price", ""))
            # Clean price formatting
            price = price.replace("€", "").replace("$", "").replace("£", "").strip()
            row[best_col_map['price']].text = price
        
        # Quantity column
        if 'quantity' in best_col_map and best_col_map['quantity'] < len(row):
            row[best_col_map['quantity']].text = str(item.get("quantity", "1"))
        
        # Total column (calculate if needed)
        if 'total' in best_col_map and best_col_map['total'] < len(row):
            try:
                price_val = float(str(item.get("unit_price", "0")).replace("€", "").replace(",", "").strip())
                qty_val = float(str(item.get("quantity", "1")))
                total_val = price_val * qty_val
                row[best_col_map['total']].text = f"{total_val:,.0f}".replace(",", " ")
            except:
                row[best_col_map['total']].text = ""
        
        # Apply basic formatting to all cells
        for cell in row:
            if len(cell.paragraphs) > 0:
                para = cell.paragraphs[0]
                if len(para.runs) > 0:
                    run = para.runs[0]
                    run.font.size = Pt(10)
        
        if idx % 10 == 0:
            print(f"  Inserted {idx} items...")
    
    except Exception as e:
        print(f"  ✗ Error on item {idx}: {str(e)}")
        continue

print(f"\n✓ Successfully inserted {len(items)} items")

# Save document
try:
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    doc.save(OUTPUT_PATH)
    file_size = os.path.getsize(OUTPUT_PATH)
    print(f"\n✓ Document saved: {OUTPUT_PATH}")
    print(f"  File size: {file_size:,} bytes")
except Exception as e:
    print(f"\n✗ Error saving: {str(e)}")
    exit(1)

print("=" * 60)
print("GENERATION COMPLETE")
print("=" * 60)