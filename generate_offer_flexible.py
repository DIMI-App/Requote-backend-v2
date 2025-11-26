import os
import sys
import json
import openai
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from collections import OrderedDict

openai.api_key = os.environ.get('OPENAI_API_KEY')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
ITEMS_PATH = os.path.join(OUTPUT_FOLDER, "items_offer1.json")

# SIMPLIFIED PROMPT 3: Just get column mapping, skip complex analysis
SIMPLE_MERGE_PROMPT = """You have extracted data from a supplier quotation. Now map it to a template format.

EXTRACTED DATA:
{extraction_json}

TEMPLATE INFO:
- Format: {template_type}
- Language: To be detected
- Has pricing table: Yes

Return ONLY this JSON (no explanations):
{{
  "column_mapping": {{
    "col_0": "position_number",
    "col_1": "description",
    "col_2": "unit_price",
    "col_3": "quantity",
    "col_4": "total_price"
  }},
  "currency_format": "€1.234,56",
  "template_language": "UK"
}}
"""

def load_extraction_data():
    """Load Offer 1 extraction JSON"""
    try:
        with open(ITEMS_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading extraction data: {e}", flush=True)
        return None

def get_simple_merge_instructions(extraction_data, template_type):
    """Get simplified merge instructions - just column mapping"""
    try:
        print("Getting simple merge instructions...", flush=True)
        
        # Simplify extraction data to avoid token limits
        simplified = {
            "items_count": len(extraction_data.get('items', [])),
            "sample_item": extraction_data.get('items', [{}])[0] if extraction_data.get('items') else {},
            "has_technical_sections": len(extraction_data.get('technical_sections', [])) > 0
        }
        
        prompt = SIMPLE_MERGE_PROMPT.format(
            extraction_json=json.dumps(simplified, ensure_ascii=False, indent=2),
            template_type=template_type
        )
        
        print("Calling GPT-4o for merge mapping...", flush=True)
        
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",  # Faster, cheaper, more reliable
            messages=[{"role": "user", "content": prompt}],
            max_tokens=500,  # Small response
            temperature=0
        )
        
        merge_json = response.choices[0].message.content.strip()
        
        # Clean JSON
        if merge_json.startswith("```json"):
            merge_json = merge_json.replace("```json", "").replace("```", "").strip()
        elif merge_json.startswith("```"):
            merge_json = merge_json.replace("```", "").strip()
        
        print(f"Raw GPT response: {merge_json[:200]}...", flush=True)
        
        merge_instructions = json.loads(merge_json)
        print("✓ Merge instructions parsed", flush=True)
        
        return merge_instructions
        
    except json.JSONDecodeError as e:
        print(f"JSON Parse Error: {e}", flush=True)
        print(f"Response was: {merge_json if 'merge_json' in locals() else 'No response'}", flush=True)
        return None
    except Exception as e:
        print(f"Error getting merge instructions: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return None

def apply_simple_merge_to_docx(template_path, extraction_data, output_path):
    """Apply extraction data to DOCX using simple column mapping"""
    try:
        print("Applying simple merge to DOCX...", flush=True)
        
        doc = Document(template_path)
        
        if not doc.tables:
            print("No tables in template", flush=True)
            return False
        
        # Find biggest table (likely pricing table)
        table = max(doc.tables, key=lambda t: len(t.rows))
        print(f"Using table with {len(table.rows)} rows, {len(table.columns)} columns", flush=True)
        
        # Clear existing data rows (keep header)
        while len(table.rows) > 1:
            table._tbl.remove(table.rows[1]._tr)
        
        # Get items
        items = extraction_data.get('items', [])
        print(f"Processing {len(items)} items", flush=True)
        
        # Group by category
        categorized = OrderedDict()
        for item in items:
            cat = item.get('category', 'Items')
            if cat not in categorized:
                categorized[cat] = []
            categorized[cat].append(item)
        
        item_counter = 1
        for category, cat_items in categorized.items():
            # Add category row
            cat_row = table.add_row().cells
            if len(cat_row) >= 2:
                cat_row[1].text = category
                for para in cat_row[1].paragraphs:
                    for run in para.runs:
                        run.font.bold = True
                        run.font.size = Pt(11)
            
            # Add item rows
            for item in cat_items:
                row = table.add_row().cells
                
                # Simple mapping
                if len(row) >= 1:
                    row[0].text = str(item_counter)
                if len(row) >= 2:
                    # Build description
                    desc_parts = []
                    if item.get('item_name'):
                        desc_parts.append(item['item_name'])
                    if item.get('technical_description'):
                        desc_parts.append(item['technical_description'])
                    if item.get('specifications'):
                        if isinstance(item['specifications'], dict):
                            specs_text = ", ".join([f"{k}: {v}" for k, v in item['specifications'].items()])
                            desc_parts.append(specs_text)
                        else:
                            desc_parts.append(str(item['specifications']))
                    
                    row[1].text = '\n\n'.join(desc_parts)
                
                if len(row) >= 3:
                    row[2].text = str(item.get('unit_price', ''))
                if len(row) >= 4:
                    row[3].text = str(item.get('quantity', '1'))
                if len(row) >= 5:
                    row[4].text = str(item.get('total_price', ''))
                
                item_counter += 1
        
        # Save
        doc.save(output_path)
        print(f"✓ DOCX saved: {output_path}", flush=True)
        return True
        
    except Exception as e:
        print(f"Error applying merge to DOCX: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return False

def generate_offer_flexible():
    """Main function - simplified version"""
    try:
        print("=== STARTING SIMPLIFIED FLEXIBLE GENERATION ===", flush=True)
        
        if not openai.api_key:
            print("ERROR: OPENAI_API_KEY not set", flush=True)
            return False
        
        # Load extraction data
        extraction_data = load_extraction_data()
        if not extraction_data:
            print("ERROR: Could not load extraction data", flush=True)
            return False
        
        print(f"✓ Loaded {len(extraction_data.get('items', []))} items", flush=True)
        
        # Determine template type
        template_docx = os.path.join(BASE_DIR, "offer2_template.docx")
        template_xlsx = os.path.join(BASE_DIR, "offer2_template.xlsx")
        
        if os.path.exists(template_docx):
            template_path = template_docx
            template_type = 'docx'
        elif os.path.exists(template_xlsx):
            template_path = template_xlsx
            template_type = 'xlsx'
        else:
            print("ERROR: No template found", flush=True)
            return False
        
        print(f"✓ Template type: {template_type}", flush=True)
        
        # Get simple merge instructions (optional - we can skip for now)
        # merge_instructions = get_simple_merge_instructions(extraction_data, template_type)
        
        # Just apply directly to DOCX
        if template_type == 'docx':
            output_path = os.path.join(OUTPUT_FOLDER, "final_offer1.docx")
            success = apply_simple_merge_to_docx(template_path, extraction_data, output_path)
        else:
            print("XLSX not yet supported in simplified version", flush=True)
            return False
        
        if success:
            print("=== SIMPLIFIED FLEXIBLE GENERATION COMPLETED ===", flush=True)
            return True
        else:
            print("ERROR: Generation failed", flush=True)
            return False
        
    except Exception as e:
        print(f"FATAL ERROR: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Simplified Flexible Generation Script Started", flush=True)
    
    success = generate_offer_flexible()
    
    if not success:
        print("Generation failed", flush=True)
        sys.exit(1)
    
    print("COMPLETED SUCCESSFULLY", flush=True)
    sys.exit(0)