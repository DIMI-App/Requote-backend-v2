import os
import json
import re
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

def detect_number_format(table):
    """Detect number format from existing template data"""
    # Scan template for existing numbers to understand format
    # Default: space-separated thousands, no decimals (324 400)
    format_info = {
        'thousands_sep': ' ',
        'decimal_sep': '',
        'decimals': 0,
        'currency_symbol': False
    }
    
    print("Analyzing template number format...", flush=True)
    
    # Scan first few rows for number patterns
    for row_idx, row in enumerate(table.rows[:5]):
        for cell in row.cells:
            text = cell.text.strip()
            # Look for numbers with separators
            if re.search(r'\d[\s.,]\d', text):
                print(f"  Sample: '{text}'", flush=True)
                
                # Detect pattern: "324 400" (space separator, no decimals)
                if re.search(r'\d+\s\d+$', text):
                    format_info['thousands_sep'] = ' '
                    format_info['decimals'] = 0
                # Detect pattern: "324.400,00" (dot thousands, comma decimals)
                elif re.search(r'\d+\.\d{3},\d{2}', text):
                    format_info['thousands_sep'] = '.'
                    format_info['decimal_sep'] = ','
                    format_info['decimals'] = 2
                # Detect pattern: "324,400.00" (comma thousands, dot decimals)
                elif re.search(r'\d+,\d{3}\.\d{2}', text):
                    format_info['thousands_sep'] = ','
                    format_info['decimal_sep'] = '.'
                    format_info['decimals'] = 2
                
                break
        if format_info['thousands_sep']:
            break
    
    print(f"✓ Detected format: thousands='{format_info['thousands_sep']}', decimals={format_info['decimals']}", flush=True)
    return format_info

def format_price(price_str, format_info):
    """Format price according to template format"""
    if not price_str:
        return ""
    
    price_lower = str(price_str).lower().strip()
    
    # Handle special states
    if 'included' in price_lower:
        return "Included"
    if any(x in price_lower for x in ['on request', 'to be quoted', 'can be offered', 'please inquire']):
        return "On request"
    
    # Extract numeric value
    # Remove currency symbols and letters
    numeric = re.sub(r'[^\d.,]', '', str(price_str))
    if not numeric:
        return price_str
    
    # Parse number (handle both formats)
    try:
        # Replace separators to get clean number
        clean = numeric.replace('.', '').replace(',', '.')
        value = float(clean)
        
        # Format according to template
        if format_info['decimals'] == 0:
            # No decimals: "324 400"
            formatted = f"{int(value):,}".replace(',', format_info['thousands_sep'])
        else:
            # With decimals: "324.400,00" or "324,400.00"
            int_part = int(value)
            dec_part = int((value - int_part) * 100)
            int_formatted = f"{int_part:,}".replace(',', format_info['thousands_sep'])
            formatted = f"{int_formatted}{format_info['decimal_sep']}{dec_part:02d}"
        
        return formatted
    except:
        # If parsing fails, return original
        return price_str

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

# Find pricing table
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

# Detect number format from template
number_format = detect_number_format(best_table)

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
    
    # Add category header row (ONLY in description column)
    category_row = best_table.add_row().cells
    
    # Clear all cells first
    for cell in category_row:
        cell.text = ""
    
    # Put category name only in description column (column 1, 0-indexed)
    if len(category_row) >= 2:
        category_row[1].text = category
        
        # Format category header cell
        for paragraph in category_row[1].paragraphs:
            paragraph.alignment = 0  # Left align
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(11)
                run.font.color.rgb = RGBColor(255, 255, 255)  # White text
        
        # Background color (blue)
        set_cell_background(category_row[1], "4472C4")
    
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
            
            # Column 2: Unit Price (formatted)
            if len(row) >= 3:
                price = item.get("unit_price", "")
                row[2].text = format_price(price, number_format)
            
            # Column 3: Quantity
            if len(row) >= 4:
                row[3].text = str(item.get("quantity", "1"))
            
            # Column 4: Total Price (formatted)
            if len(row) >= 5:
                total = item.get("total_price", "")
                row[4].text = format_price(total, number_format)
            
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