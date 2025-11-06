import os
import json
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from collections import Counter
import openai

openai.api_key = os.environ.get('OPENAI_API_KEY')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OFFER_2_PATH = os.path.join(BASE_DIR, "offer2_template.docx")
ITEMS_PATH = os.path.join(BASE_DIR, "outputs", "items_offer1.json")
OUTPUT_PATH = os.path.join(BASE_DIR, "outputs", "final_offer1.docx")

def detect_template_language(doc):
    """Detect the language of Offer 2 template"""
    print("=" * 60, flush=True)
    print("DETECTING TEMPLATE LANGUAGE", flush=True)
    print("=" * 60, flush=True)
    
    # Extract text samples from template
    text_samples = []
    
    # Get text from tables
    for table in doc.tables:
        for row in table.rows[:5]:  # First 5 rows
            for cell in row.cells:
                text = cell.text.strip()
                if len(text) > 10:  # Meaningful text
                    text_samples.append(text)
    
    # Get text from paragraphs
    for para in doc.paragraphs[:10]:  # First 10 paragraphs
        text = para.text.strip()
        if len(text) > 10:
            text_samples.append(text)
    
    if not text_samples:
        print("⚠ No text found in template, defaulting to English", flush=True)
        return "en"
    
    # Combine samples
    combined_text = " ".join(text_samples[:10])  # First 10 samples
    print(f"Text sample for detection: {combined_text[:200]}...", flush=True)
    
    # Use OpenAI to detect language
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{
                "role": "user",
                "content": f"""Detect the language of this text. Return ONLY the two-letter ISO 639-1 language code (e.g., 'en', 'uk', 'es', 'de', 'fr', 'it', 'ru', 'pl').

Text: {combined_text[:500]}

Language code:"""
            }],
            max_tokens=10,
            temperature=0
        )
        
        detected_lang = response.choices[0].message.content.strip().lower()
        
        # Validate
        valid_codes = ['en', 'uk', 'es', 'de', 'fr', 'it', 'ru', 'pl', 'pt', 'nl', 'tr', 'ar', 'zh', 'ja', 'ko']
        if detected_lang not in valid_codes:
            detected_lang = 'en'
        
        lang_names = {
            'en': 'English', 'uk': 'Ukrainian', 'es': 'Spanish', 'de': 'German',
            'fr': 'French', 'it': 'Italian', 'ru': 'Russian', 'pl': 'Polish',
            'pt': 'Portuguese', 'nl': 'Dutch', 'tr': 'Turkish', 'ar': 'Arabic',
            'zh': 'Chinese', 'ja': 'Japanese', 'ko': 'Korean'
        }
        
        print(f"✓ Detected language: {lang_names.get(detected_lang, detected_lang)} ({detected_lang})", flush=True)
        print("=" * 60, flush=True)
        
        return detected_lang
        
    except Exception as e:
        print(f"✗ Language detection error: {str(e)}", flush=True)
        print("Defaulting to English", flush=True)
        return "en"

def translate_items(items, target_lang):
    """Translate all items to target language"""
    if target_lang == 'en':
        print("Target language is English, no translation needed", flush=True)
        return items
    
    print("=" * 60, flush=True)
    print(f"TRANSLATING ITEMS TO {target_lang.upper()}", flush=True)
    print("=" * 60, flush=True)
    
    try:
        # Prepare items for translation
        items_to_translate = []
        for item in items:
            items_to_translate.append({
                'category': item.get('category', ''),
                'item_name': item.get('item_name', ''),
                'details': item.get('details', '')
            })
        
        # Create translation request
        translation_prompt = f"""Translate the following product quotation items to {target_lang.upper()}. 

IMPORTANT RULES:
1. Translate category names, item names, and details
2. Keep technical terms and model numbers in ORIGINAL language (e.g., "CAN ISO 20/2 S", "VBS MINIDOSE")
3. Keep measurements, numbers, and units in original format
4. Preserve proper nouns, brand names, and technical specifications
5. Return ONLY valid JSON array with same structure

Input JSON:
{json.dumps(items_to_translate, ensure_ascii=False, indent=2)}

Output (translated JSON):"""
        
        print("Sending translation request to OpenAI...", flush=True)
        
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{
                "role": "user",
                "content": translation_prompt
            }],
            max_tokens=6000,
            temperature=0.3
        )
        
        translated_json = response.choices[0].message.content.strip()
        
        # Clean JSON
        if translated_json.startswith("```json"):
            translated_json = translated_json.replace("```json", "").replace("```", "").strip()
        elif translated_json.startswith("```"):
            translated_json = translated_json.replace("```", "").strip()
        
        translated_items = json.loads(translated_json)
        
        print(f"✓ Translated {len(translated_items)} items", flush=True)
        
        # Merge translated text back into original items (preserve prices/quantities)
        for i, item in enumerate(items):
            if i < len(translated_items):
                item['category'] = translated_items[i].get('category', item['category'])
                item['item_name'] = translated_items[i].get('item_name', item['item_name'])
                item['details'] = translated_items[i].get('details', item['details'])
        
        print("=" * 60, flush=True)
        return items
        
    except Exception as e:
        print(f"✗ Translation error: {str(e)}", flush=True)
        print("Continuing with original language", flush=True)
        import traceback
        traceback.print_exc()
        return items

# [Keep all the existing functions: get_cell_background_color, get_text_color, etc.]
def get_cell_background_color(cell):
    """Extract background color from cell"""
    try:
        tcPr = cell._element.get_or_add_tcPr()
        shd = tcPr.find(qn('w:shd'))
        if shd is not None:
            fill = shd.get(qn('w:fill'))
            if fill and fill != 'auto':
                return fill
    except:
        pass
    return None

def get_text_color(run):
    """Extract text color from run"""
    try:
        if run.font.color.rgb:
            return str(run.font.color.rgb)
    except:
        pass
    return None

def get_font_name(run):
    """Extract font name from run"""
    try:
        if run.font.name:
            return run.font.name
    except:
        pass
    return None

def get_font_size(run):
    """Extract font size from run"""
    try:
        if run.font.size:
            return run.font.size.pt
    except:
        pass
    return None

def analyze_template_style(doc):
    """Analyze template to detect colors and fonts"""
    style_info = {
        'header_bg_color': None,
        'header_text_color': None,
        'body_text_color': None,
        'primary_font': None,
        'header_font_size': 11,
        'body_font_size': 10
    }
    
    print("=" * 60, flush=True)
    print("ANALYZING TEMPLATE STYLE", flush=True)
    print("=" * 60, flush=True)
    
    bg_colors = []
    text_colors = []
    fonts = []
    font_sizes = []
    
    for table in doc.tables:
        for row_idx, row in enumerate(table.rows):
            for cell in row.cells:
                bg = get_cell_background_color(cell)
                if bg:
                    bg_colors.append(bg)
                
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        color = get_text_color(run)
                        if color:
                            text_colors.append(color)
                        
                        font = get_font_name(run)
                        if font:
                            fonts.append(font)
                        
                        size = get_font_size(run)
                        if size:
                            font_sizes.append(size)
    
    if bg_colors:
        bg_counter = Counter(bg_colors)
        style_info['header_bg_color'] = bg_counter.most_common(1)[0][0]
        print(f"✓ Header background: #{style_info['header_bg_color']}", flush=True)
    
    if text_colors:
        text_counter = Counter(text_colors)
        most_common = text_counter.most_common(2)
        for color, count in most_common:
            if color.upper() in ['FFFFFF', 'FFFFFFFF']:
                style_info['header_text_color'] = color
                print(f"✓ Header text: #{color}", flush=True)
            else:
                style_info['body_text_color'] = color
                print(f"✓ Body text: #{color}", flush=True)
    
    if fonts:
        font_counter = Counter(fonts)
        style_info['primary_font'] = font_counter.most_common(1)[0][0]
        print(f"✓ Primary font: {style_info['primary_font']}", flush=True)
    
    if font_sizes:
        size_counter = Counter(font_sizes)
        common_sizes = size_counter.most_common(2)
        if len(common_sizes) >= 2:
            sizes_sorted = sorted([s[0] for s in common_sizes], reverse=True)
            style_info['header_font_size'] = sizes_sorted[0]
            style_info['body_font_size'] = sizes_sorted[1]
        elif len(common_sizes) == 1:
            style_info['body_font_size'] = common_sizes[0][0]
        
        print(f"✓ Header font size: {style_info['header_font_size']}pt", flush=True)
        print(f"✓ Body font size: {style_info['body_font_size']}pt", flush=True)
    
    print("=" * 60, flush=True)
    
    return style_info

def set_cell_background(cell, color_hex):
    """Set cell background color"""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), color_hex)
    cell._element.get_or_add_tcPr().append(shading_elm)

def apply_text_style(cell, text, is_header, style_info):
    """Apply font and color styling to cell"""
    cell.text = text
    
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            if style_info['primary_font']:
                run.font.name = style_info['primary_font']
            
            if is_header:
                run.font.size = Pt(style_info['header_font_size'])
                run.bold = True
                if style_info['header_text_color']:
                    color_hex = style_info['header_text_color'].replace('#', '')
                    if len(color_hex) == 6:
                        run.font.color.rgb = RGBColor(
                            int(color_hex[0:2], 16),
                            int(color_hex[2:4], 16),
                            int(color_hex[4:6], 16)
                        )
            else:
                run.font.size = Pt(style_info['body_font_size'])
                if style_info['body_text_color']:
                    color_hex = style_info['body_text_color'].replace('#', '')
                    if len(color_hex) == 6:
                        run.font.color.rgb = RGBColor(
                            int(color_hex[0:2], 16),
                            int(color_hex[2:4], 16),
                            int(color_hex[4:6], 16)
                        )

def detect_number_format(table):
    """Detect number format from existing template data"""
    format_info = {
        'thousands_sep': ' ',
        'decimal_sep': '',
        'decimals': 0,
        'currency_symbol': False
    }
    
    print("Analyzing template number format...", flush=True)
    
    for row_idx, row in enumerate(table.rows[:5]):
        for cell in row.cells:
            text = cell.text.strip()
            if re.search(r'\d[\s.,]\d', text):
                print(f"  Sample: '{text}'", flush=True)
                
                if re.search(r'\d+\s\d+$', text):
                    format_info['thousands_sep'] = ' '
                    format_info['decimals'] = 0
                elif re.search(r'\d+\.\d{3},\d{2}', text):
                    format_info['thousands_sep'] = '.'
                    format_info['decimal_sep'] = ','
                    format_info['decimals'] = 2
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
    
    if 'included' in price_lower:
        return "Included"
    if any(x in price_lower for x in ['on request', 'to be quoted', 'can be offered', 'please inquire']):
        return "On request"
    
    numeric = re.sub(r'[^\d.,]', '', str(price_str))
    if not numeric:
        return price_str
    
    try:
        clean = numeric.replace('.', '').replace(',', '.')
        value = float(clean)
        
        if format_info['decimals'] == 0:
            formatted = f"{int(value):,}".replace(',', format_info['thousands_sep'])
        else:
            int_part = int(value)
            dec_part = int((value - int_part) * 100)
            int_formatted = f"{int_part:,}".replace(',', format_info['thousands_sep'])
            formatted = f"{int_formatted}{format_info['decimal_sep']}{dec_part:02d}"
        
        return formatted
    except:
        return price_str

# MAIN EXECUTION
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

# Detect template language
target_language = detect_template_language(doc)

# Translate items to target language
items = translate_items(items, target_language)

# Analyze template style
template_style = analyze_template_style(doc)

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

# Detect number format
number_format = detect_number_format(best_table)

# Clear existing data rows
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

# Insert items
item_counter = 1
for category, cat_items in categorized_items.items():
    print(f"  Processing category: {category} ({len(cat_items)} items)", flush=True)
    
    category_row = best_table.add_row().cells
    
    for cell in category_row:
        cell.text = ""
    
    if len(category_row) >= 2:
        apply_text_style(category_row[1], category, True, template_style)
        
        if template_style['header_bg_color']:
            set_cell_background(category_row[1], template_style['header_bg_color'])
    
    for item in cat_items:
        row = best_table.add_row().cells
        
        try:
            if len(row) >= 1:
                apply_text_style(row[0], f"{item_counter}.", False, template_style)
                item_counter += 1
            
            if len(row) >= 2:
                desc = item.get("item_name", "")
                if item.get("details"):
                    desc = f"{desc}\n{item.get('details')}"
                apply_text_style(row[1], desc, False, template_style)
            
            if len(row) >= 3:
                price = format_price(item.get("unit_price", ""), number_format)
                apply_text_style(row[2], price, False, template_style)
            
            if len(row) >= 4:
                apply_text_style(row[3], str(item.get("quantity", "1")), False, template_style)
            
            if len(row) >= 5:
                total = format_price(item.get("total_price", ""), number_format)
                apply_text_style(row[4], total, False, template_style)
            
        except Exception as e:
            print(f"  ✗ Error on item: {str(e)}", flush=True)
            continue

print(f"✓ Inserted all items with template styling", flush=True)

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