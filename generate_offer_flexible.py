
import os
import sys
import json
import openai
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import openpyxl
from openpyxl.styles import Font, Alignment
from collections import OrderedDict

openai.api_key = os.environ.get('OPENAI_API_KEY')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
ITEMS_PATH = os.path.join(OUTPUT_FOLDER, "items_offer1.json")
TEMPLATE_STRUCTURE_PATH = os.path.join(OUTPUT_FOLDER, "template_structure.json")

# PROMPT 3: Merge and Create Offer 3
MERGE_PROMPT_START = """PROMPT 3: MERGE TWO JSONS AND CREATE OFFER 3
==============================================

You are a skilled document specialist who creates professional quotations by merging supplier data into company templates. You've just received two JSON files:

1. **EXTRACTED DATA (from Offer 1)** - All items, prices, descriptions, specifications, and images from supplier
2. **TEMPLATE STRUCTURE (from Offer 2)** - Your company's branded template with exact formatting rules

YOUR MISSION:
Create a new quotation (Offer 3) that looks EXACTLY like your company template (Offer 2) but contains ALL the data from the supplier (Offer 1). This should look like a human manually retyped everything from Offer 1 into Offer 2.

CRITICAL RULES FOR MERGING:

1. **PRESERVE TEMPLATE STRUCTURE 100%**
   - Keep exact section order from Offer 2
   - Maintain all branding elements (logo, colors, fonts)
   - Use same column structure and table format
   - Follow template's layout flow exactly
   - Keep header/footer unchanged

2. **INSERT ALL DATA FROM OFFER 1**
   - Every single item with pricing
   - All technical descriptions
   - Complete specifications
   - All images
   - Every optional/accessory item
   - Nothing should be left out

3. **INTELLIGENT CONTENT PLACEMENT**
   Based on Offer 2 structure, place Offer 1 content intelligently.

Return a detailed JSON structure that specifies EXACTLY how to populate the template with the extracted data.

NOW MERGE THESE TWO JSONS:

**JSON 1 - EXTRACTED FROM OFFER 1 (Supplier Data):**
"""

MERGE_PROMPT_END = """

**JSON 2 - TEMPLATE STRUCTURE FROM OFFER 2:**
"""

MERGE_PROMPT_FINAL = """

CREATE OFFER 3 STRUCTURE - Return JSON with these keys:
{
  "column_mapping": {
    "description": "Map Offer 1 fields to Offer 2 template columns",
    "template_column_1": "Offer_1_field_name",
    "template_column_2": "Offer_1_field_name",
    ...
  },
  "content_placement": {
    "where_to_put_descriptions": "in_table_column / before_table / after_table",
    "where_to_put_images": "after_table / inline / separate_section",
    "where_to_put_technical_specs": "in_description / separate_section / footnotes"
  },
  "formatting_rules": {
    "currency_format": "from template analysis",
    "number_format": "from template analysis",
    "category_style": "how to show categories in template"
  },
  "item_mapping": [
    {
      "offer1_item_index": 0,
      "category": "from Offer 1",
      "map_to_template_row": "how to fill this row",
      "column_values": {
        "col1": "value",
        "col2": "value",
        ...
      }
    }
  ],
  "generation_instructions": "Step-by-step guide for populating template"
}
"""

def load_extraction_data():
    """Load Offer 1 extraction JSON"""
    try:
        with open(ITEMS_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading extraction data: {e}", flush=True)
        return None

def load_template_structure():
    """Load Offer 2 template analysis JSON"""
    try:
        with open(TEMPLATE_STRUCTURE_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading template structure: {e}", flush=True)
        return None

def get_merge_instructions(extraction_data, template_structure):
    """Use PROMPT 3 to get intelligent merge instructions"""
    try:
        print("=== GENERATING MERGE INSTRUCTIONS (PROMPT 3) ===", flush=True)
        
        # Build full prompt
        full_prompt = MERGE_PROMPT_START
        full_prompt += json.dumps(extraction_data, indent=2, ensure_ascii=False)
        full_prompt += MERGE_PROMPT_END
        full_prompt += json.dumps(template_structure, indent=2, ensure_ascii=False)
        full_prompt += MERGE_PROMPT_FINAL
        
        print("Calling GPT-4o for merge instructions...", flush=True)
        
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": full_prompt}],
            max_tokens=6000,
            temperature=0
        )
        
        print("Received merge instructions", flush=True)
        
        merge_json = response.choices[0].message.content.strip()
        
        # Clean JSON formatting
        if merge_json.startswith("```json"):
            merge_json = merge_json.replace("```json", "").replace("```", "").strip()
        elif merge_json.startswith("```"):
            merge_json = merge_json.replace("```", "").strip()
        
        merge_instructions = json.loads(merge_json)
        
        return merge_instructions
        
    except Exception as e:
        print(f"Error getting merge instructions: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return None

def apply_merge_to_docx(template_path, extraction_data, merge_instructions, output_path):
    """Apply merge instructions to DOCX template"""
    try:
        print("Applying merge to DOCX template...", flush=True)
        
        doc = Document(template_path)
        
        # Find pricing table based on merge instructions
        if 'pricing_table' in merge_instructions:
            table_location = merge_instructions['pricing_table'].get('location', 'main_table')
            print(f"Looking for pricing table: {table_location}", flush=True)
        
        # Get column mapping from merge instructions
        column_mapping = merge_instructions.get('column_mapping', {})
        print(f"Column mapping: {column_mapping}", flush=True)
        
        # Find the table
        if not doc.tables:
            print("No tables found in template", flush=True)
            return False
        
        table = doc.tables[0]  # Use first table for now
        print(f"Found table with {len(table.rows)} rows, {len(table.columns)} columns", flush=True)
        
        # Clear existing data rows (keep header)
        while len(table.rows) > 1:
            table._tbl.remove(table.rows[1]._tr)
        
        # Get items and apply mapping
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
            
            # Add item rows
            for item in cat_items:
                row = table.add_row().cells
                
                # Apply column mapping from merge instructions
                if len(row) >= 1:
                    row[0].text = str(item_counter)
                if len(row) >= 2:
                    # Build description based on merge instructions
                    desc_parts = []
                    if item.get('item_name'):
                        desc_parts.append(item['item_name'])
                    if item.get('technical_description'):
                        desc_parts.append(item['technical_description'])
                    row[1].text = '\n\n'.join(desc_parts)
                if len(row) >= 3:
                    row[2].text = str(item.get('unit_price', ''))
                if len(row) >= 4:
                    row[3].text = str(item.get('quantity', '1'))
                if len(row) >= 5:
                    row[4].text = str(item.get('total_price', ''))
                
                item_counter += 1
        
        # Save document
        doc.save(output_path)
        print(f"✓ DOCX saved: {output_path}", flush=True)
        return True
        
    except Exception as e:
        print(f"Error applying merge to DOCX: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return False

def apply_merge_to_xlsx(template_path, extraction_data, merge_instructions, output_path):
    """Apply merge instructions to XLSX template"""
    try:
        print("Applying merge to XLSX template...", flush=True)
        
        workbook = openpyxl.load_workbook(template_path)
        sheet = workbook.active
        
        # Find table header based on merge instructions
        table_start_row = None
        for row_idx in range(1, min(20, sheet.max_row + 1)):
            row = sheet[row_idx]
            row_text = ' '.join([str(cell.value).upper() if cell.value else '' for cell in row])
            if any(kw in row_text for kw in ['DESCRIPTION', 'PRICE', 'QUANTITY']):
                table_start_row = row_idx
                break
        
        if not table_start_row:
            print("Could not find table header", flush=True)
            return False
        
        print(f"Found table at row {table_start_row}", flush=True)
        
        # Delete existing data
        data_start = table_start_row + 1
        if sheet.max_row >= data_start:
            sheet.delete_rows(data_start, sheet.max_row - table_start_row)
        
        # Get column mapping
        column_mapping = merge_instructions.get('column_mapping', {})
        
        # Insert items
        items = extraction_data.get('items', [])
        current_row = data_start
        
        for item in items:
            sheet.insert_rows(current_row)
            
            # Apply mapping (simplified - would use merge instructions in production)
            sheet.cell(current_row, 1).value = item.get('item_name', '')
            sheet.cell(current_row, 2).value = item.get('quantity', '1')
            sheet.cell(current_row, 3).value = item.get('unit_price', '')
            sheet.cell(current_row, 4).value = item.get('total_price', '')
            
            current_row += 1
        
        # Save workbook
        workbook.save(output_path)
        print(f"✓ XLSX saved: {output_path}", flush=True)
        return True
        
    except Exception as e:
        print(f"Error applying merge to XLSX: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return False

def generate_offer_flexible():
    """Main function using PROMPT 3 approach"""
    try:
        print("=== STARTING FLEXIBLE OFFER GENERATION (PROMPT 3) ===", flush=True)
        
        if not openai.api_key:
            print("ERROR: OPENAI_API_KEY not set", flush=True)
            return False
        
        # Load extraction data (Offer 1)
        extraction_data = load_extraction_data()
        if not extraction_data:
            print("ERROR: Could not load extraction data", flush=True)
            return False
        
        print(f"✓ Loaded extraction data: {len(extraction_data.get('items', []))} items", flush=True)
        
        # Load template structure (Offer 2)
        template_structure = load_template_structure()
        if not template_structure:
            print("ERROR: Could not load template structure", flush=True)
            return False
        
        print("✓ Loaded template structure", flush=True)
        
        # Get merge instructions using PROMPT 3
        merge_instructions = get_merge_instructions(extraction_data, template_structure)
        if not merge_instructions:
            print("ERROR: Could not get merge instructions", flush=True)
            return False
        
        print("✓ Generated merge instructions", flush=True)
        
        # Save merge instructions for debugging
        merge_path = os.path.join(OUTPUT_FOLDER, "merge_instructions.json")
        with open(merge_path, 'w', encoding='utf-8') as f:
            json.dump(merge_instructions, f, indent=2, ensure_ascii=False)
        print(f"✓ Saved merge instructions to {merge_path}", flush=True)
        
        # Determine template type and apply merge
        template_type = template_structure.get('template_type', 'docx')
        
        if template_type == 'docx':
            template_path = os.path.join(BASE_DIR, "offer2_template.docx")
            output_path = os.path.join(OUTPUT_FOLDER, "final_offer1.docx")
            success = apply_merge_to_docx(template_path, extraction_data, merge_instructions, output_path)
        elif template_type in ['xlsx', 'xls']:
            template_path = os.path.join(BASE_DIR, "offer2_template.xlsx")
            output_path = os.path.join(OUTPUT_FOLDER, "final_offer1.xlsx")
            success = apply_merge_to_xlsx(template_path, extraction_data, merge_instructions, output_path)
        else:
            print(f"Unsupported template type: {template_type}", flush=True)
            return False
        
        if success:
            print("=== FLEXIBLE OFFER GENERATION COMPLETED ===", flush=True)
            return True
        else:
            print("ERROR: Merge application failed", flush=True)
            return False
        
    except Exception as e:
        print(f"FATAL ERROR: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Flexible Offer Generation Script Started (PROMPT 3)", flush=True)
    
    success = generate_offer_flexible()
    
    if not success:
        print("Flexible generation failed", flush=True)
        sys.exit(1)
    
    print("COMPLETED SUCCESSFULLY (PROMPT 3)", flush=True)
    sys.exit(0)