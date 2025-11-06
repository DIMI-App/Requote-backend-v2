import os
import json
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from collections import Counter
import openai
import deepl

openai.api_key = os.environ.get('OPENAI_API_KEY')
deepl_key = os.environ.get('DEEPL_API_KEY')

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
        for row in table.rows[:5]:
            for cell in row.cells:
                text = cell.text.strip()
                if len(text) > 10:
                    text_samples.append(text)
    
    # Get text from paragraphs
    for para in doc.paragraphs[:10]:
        text = para.text.strip()
        if len(text) > 10:
            text_samples.append(text)
    
    if not text_samples:
        print("⚠ No text found in template, defaulting to English", flush=True)
        return "EN"
    
    # Combine samples
    combined_text = " ".join(text_samples[:5])[:500]
    print(f"Text sample: {combined_text[:150]}...", flush=True)
    
    # Use DeepL to detect language
    try:
        if not deepl_key:
            print("⚠ No DeepL API key, using OpenAI for detection", flush=True)
            return detect_language_openai(combined_text)
        
        translator = deepl.Translator(deepl_key)
        result = translator.translate_text(combined_text, target_lang="EN-US")
        
        detected_lang = result.detected_source_lang
        
        print(f"✓ Detected language: {detected_lang}", flush=True)
        print("=" * 60, flush=True)
        
        return detected_lang
        
    except Exception as e:
        print(f"✗ DeepL detection error: {str(e)}", flush=True)
        print("Falling back to OpenAI detection", flush=True)
        return detect_language_openai(combined_text)

def detect_language_openai(text):
    """Fallback language detection using OpenAI"""
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{
                "role": "user",
                "content": f"""Detect the language of this text. Return ONLY the two-letter ISO 639-1 code in UPPERCASE (e.g., 'EN', 'UK', 'ES', 'DE').

Text: {text}

Language code:"""
            }],
            max_tokens=10,
            temperature=0
        )
        
        detected_lang = response.choices[0].message.content.strip().upper()
        return detected_lang if len(detected_lang) == 2 else "EN"
        
    except Exception as e:
        print(f"✗ OpenAI detection error: {str(e)}", flush=True)
        return "EN"

def translate_with_deepl(text, target_lang, preserve_formatting=False):
    """Translate text using DeepL API"""
    try:
        if not deepl_key:
            raise Exception("DeepL API key not configured")
        
        translator = deepl.Translator(deepl_key)
        
        # DeepL language codes (uppercase)
        # Handle special cases
        if target_lang == "UK":
            target_lang = "UK"  # Ukrainian
        elif target_lang == "EN":
            target_lang = "EN-US"
        
        result = translator.translate_text(
            text,
            target_lang=target_lang,
            preserve_formatting=preserve_formatting,
            tag_handling="xml" if preserve_formatting else None
        )
        
        return result.text
        
    except Exception as e:
        print(f"DeepL translation error: {str(e)}", flush=True)
        raise

def translate_items(items, target_lang):
    """Translate all items to target language using DeepL"""
    if target_lang == 'EN' or target_lang == 'EN-US':
        print("Target language is English, no translation needed", flush=True)
        return items
    
    print("=" * 60, flush=True)
    print(f"TRANSLATING ITEMS TO {target_lang} USING DEEPL", flush=True)
    print("=" * 60, flush=True)
    
    if not deepl_key:
        print("✗ DeepL API key not configured", flush=True)
        print("Falling back to OpenAI translation", flush=True)
        return translate_items_openai(items, target_lang)
    
    try:
        translator = deepl.Translator(deepl_key)
        
        # Get unique categories
        categories = list(set([item.get('category', '') for item in items if item.get('category')]))
        print(f"Translating {len(categories)} categories...", flush=True)
        
        # Translate categories
        category_mapping = {}
        for category in categories:
            try:
                translated = translator.translate_text(
                    category,
                    target_lang=target_lang
                )
                category_mapping[category] = translated.text
                print(f"  '{category}' → '{translated.text}'", flush=True)
            except Exception as e:
                print(f"  ✗ Failed to translate category '{category}': {e}", flush=True)
                category_mapping[category] = category
        
        # Translate items in batches (to preserve context)
        print(f"Translating {len(items)} items...", flush=True)
        translated_items = []
        
        batch_size = 10
        for i in range(0, len(items), batch_size):
            batch = items[i:i+batch_size]
            
            # Prepare texts for batch translation
            item_names = [item.get('item_name', '') for item in batch]
            item_details = [item.get('details', '') for item in batch]
            
            # Filter out empty strings
            names_to_translate = [name for name in item_names if name]
            details_to_translate = [detail for detail in item_details if detail]
            
            try:
                # Translate item names
                if names_to_translate:
                    translated_names_result = translator.translate_text(
                        names_to_translate,
                        target_lang=target_lang
                    )
                    # Handle single vs multiple results
                    if isinstance(translated_names_result, list):
                        translated_names = [r.text for r in translated_names_result]
                    else:
                        translated_names = [translated_names_result.text]
                else:
                    translated_names = []
                
                # Translate details
                if details_to_translate:
                    translated_details_result = translator.translate_text(
                        details_to_translate,
                        target_lang=target_lang
                    )
                    if isinstance(translated_details_result, list):
                        translated_details = [r.text for r in translated_details_result]
                    else:
                        translated_details = [translated_details_result.text]
                else:
                    translated_details = []
                
                # Apply translations to batch
                name_idx = 0
                detail_idx = 0
                for item in batch:
                    translated_item = item.copy()
                    
                    # Apply category translation
                    if item.get('category') in category_mapping:
                        translated_item['category'] = category_mapping[item['category']]
                    
                    # Apply item name translation
                    if item.get('item_name'):
                        translated_item['item_name'] = translated_names[name_idx] if name_idx < len(translated_names) else item['item_name']
                        name_idx += 1
                    
                    # Apply details translation
                    if item.get('details'):
                        translated_item['details'] = translated_details[detail_idx] if detail_idx < len(translated_details) else item['details']
                        detail_idx += 1
                    
                    translated_items.append(translated_item)
                
                print(f"  Batch {i//batch_size + 1}/{(len(items)-1)//batch_size + 1} completed", flush=True)
                
            except Exception as e:
                print(f"  ✗ Batch translation failed: {e}", flush=True)
                # Keep original for this batch
                translated_items.extend(batch)
        
        # Verify translation
        if len(translated_items) > 0:
            sample_orig = items[0]
            sample_trans = translated_items[0]
            print("\nSample translation:", flush=True)
            print(f"  Category: '{sample_orig.get('category', '')}' → '{sample_trans.get('category', '')}'", flush=True)
            print(f"  Item: '{sample_orig.get('item_name', '')[:50]}' → '{sample_trans.get('item_name', '')[:50]}'", flush=True)
        
        print(f"✓ DeepL translation completed: {len(translated_items)} items", flush=True)
        print("=" * 60, flush=True)
        
        return translated_items
        
    except Exception as e:
        print(f"✗ DeepL translation failed: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        print("Falling back to OpenAI translation", flush=True)
        return translate_items_openai(items, target_lang)

def translate_items_openai(items, target_lang):
    """Fallback translation using OpenAI"""
    print("Using OpenAI for translation...", flush=True)
    
    lang_map = {
        'UK': 'Ukrainian',
        'ES': 'Spanish',
        'DE': 'German',
        'FR': 'French',
        'IT': 'Italian',
        'RU': 'Russian',
        'PL': 'Polish'
    }
    
    target_language_name = lang_map.get(target_lang, target_lang)
    
    try:
        translation_prompt = f"""Translate these quotation items to {target_language_name}. Keep technical terms and model numbers unchanged.

{json.dumps(items, ensure_ascii=False, indent=2)}

Return translated JSON:"""
        
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{
                "role": "system",
                "content": f"You are a professional translator. Translate to {target_language_name}, preserving technical terms."
            }, {
                "role": "user",
                "content": translation_prompt
            }],
            max_tokens=8000,
            temperature=0.1
        )
        
        translated_json = response.choices[0].message.content.strip()
        
        if translated_json.startswith("```"):
            translated_json = translated_json.replace("```json", "").replace("```", "").strip()
        
        return json.loads(translated_json)
        
    except Exception as e:
        print(f"✗ OpenAI translation also failed: {e}", flush=True)
        return items

# [Keep all existing style functions - get_cell_background_color, analyze_template_style, etc.]
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
print("GENERATE OFFER - SV12 with DeepL", flush=True)
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

# Translate items to target language using DeepL
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