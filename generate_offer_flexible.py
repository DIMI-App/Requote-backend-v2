import os
import sys
import json
import subprocess
import openai
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from collections import OrderedDict

openai.api_key = os.environ.get('OPENAI_API_KEY')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
ITEMS_PATH = os.path.join(OUTPUT_FOLDER, "items_offer1.json")
TEMPLATE_STRUCTURE_PATH = os.path.join(OUTPUT_FOLDER, "template_structure.json")

# REVISED PROMPT 3: INTELLIGENT DOCUMENT RECOMPOSITION  
RECOMPOSITION_PROMPT = """You are an expert sales manager who needs to rebrand a supplier's quotation into your company's format.

**YOUR TASK:**
Take a supplier's complete offer (Offer 1 - competitor/supplier product with their branding) and rewrite it as if YOUR company (Offer 2 template brand) is selling it, following YOUR company's document structure and branding.

**REAL-WORLD EXAMPLE:**
- Supplier sent you a quotation for "Cadalpe C27 Distillation Unit" with Italian branding
- You need to quote this SAME equipment but as "Robino Galandrino" offer
- Keep ALL technical content from Cadalpe
- Use Robino Galandrino's template structure and branding
- Result: looks like R&G is selling the distillation unit

**WHAT YOU HAVE:**

1. **EXTRACTED CONTENT from Offer 1 (Supplier):**
{extraction_json}

2. **TEMPLATE STRUCTURE from Offer 2 (Your Company):**
{template_structure_json}

**YOUR JOB - Act Like a Human:**

1. **UNDERSTAND THE PRODUCT**
   - What is being sold? (equipment type, model, capacity)
   - What are the key technical features?
   - What's included vs optional?

2. **MAP CONTENT TO TEMPLATE STRUCTURE**
   Following the template's `document_flow` and `content_placement`:
   
   - **If template has intro text before table:**
     → Create product overview paragraph from supplier's technical content
   
   - **If template puts descriptions before table as paragraphs:**
     → Extract technical description from items and format as paragraphs
   
   - **If template puts descriptions IN table column:**
     → Keep full technical descriptions in table cells
   
   - **If template has separate technical specs section:**
     → Extract all specifications and format as technical data table/list
   
   - **If template shows images:**
     → Note which images relate to which items
   
   - **If template has optional items section:**
     → Separate main items from optional/accessories

3. **PRESERVE TECHNICAL ACCURACY**
   - Keep ALL model numbers exactly as in Offer 1
   - Keep ALL technical specifications (dimensions, capacity, power, etc.)
   - Keep ALL pricing (adjust currency format to match template)
   - Keep ALL included/excluded items lists

4. **ADAPT TO TEMPLATE LANGUAGE**
   - If template is in Ukrainian and supplier is in English → note that translation is needed
   - If both same language → keep technical terms in original language

5. **CREATE RECOMPOSITION PLAN**

Return detailed JSON that tells the generation script EXACTLY how to rebuild the document:

{
  "product_summary": "Brief description of what's being quoted",
  
  "document_sections": [
    {
      "section_type": "intro_text",
      "position": "before_pricing_table",
      "content_source": "synthesize from items technical_descriptions",
      "formatting": "paragraph, 11pt, product overview style"
    },
    {
      "section_type": "technical_description",
      "position": "before_pricing_table",
      "content_source": "extract technical_sections from offer1",
      "formatting": "bullet_list or paragraphs based on template"
    },
    {
      "section_type": "pricing_table",
      "position": "main_body",
      "content": {
        "categories": ["Main Equipment", "Packing"],
        "column_mapping": {
          "col_0": "position_number",
          "col_1": "full_description_with_specs",
          "col_2": "unit_price",
          "col_3": "quantity",
          "col_4": "total_price"
        },
        "items_to_include": "all items from extraction"
      }
    },
    {
      "section_type": "technical_data_table",
      "position": "after_pricing_table",
      "content_source": "extract specifications from all items, create consolidated tech data table",
      "formatting": "table with NAME | UNIT | VALUE columns"
    },
    {
      "section_type": "exclusions",
      "position": "after_technical_data",
      "content_source": "notes fields marked as excluded or not included",
      "formatting": "bullet_list"
    }
  ],
  
  "content_mapping": {
    "header_company_name": "from template structure",
    "product_title": "synthesized from items",
    "technical_descriptions_placement": "before_table / in_table / after_table / separate_section",
    "images_placement": "after_table / inline / not_included",
    "optional_items_handling": "separate_section / marked_in_main_table"
  },
  
  "formatting_instructions": {
    "currency_format": "from template",
    "number_format": "from template",
    "section_headers_style": "from template formatting_guide",
    "table_style": "from template"
  },
  
  "translation_needed": true/false,
  "target_language": "UK/EN/IT/ES/DE/FR"
}

**CRITICAL RULES:**
- You're NOT just filling a form - you're RECOMPOSING a complete technical document
- Every piece of Offer 1's technical content must find its place in Offer 2's structure
- Nothing should be lost in translation
- The result must look native to Offer 2's company, not like a copy-paste job

Now create the recomposition plan:
"""

def load_extraction_data():
    """Load Offer 1 extraction"""
    try:
        with open(ITEMS_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading extraction: {e}", flush=True)
        return None

def load_template_structure():
    """Load Offer 2 template analysis"""
    try:
        with open(TEMPLATE_STRUCTURE_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return data.get('structure', {})
    except Exception as e:
        print(f"Template structure not found: {e}", flush=True)
        return None

def run_prompt_2_analysis():
    """Run PROMPT 2 to analyze template structure"""
    try:
        print("=" * 60, flush=True)
        print("RUNNING PROMPT 2 (Template Analysis)", flush=True)
        print("=" * 60, flush=True)
        
        analyze_script = os.path.join(BASE_DIR, 'analyze_offer2_template.py')
        
        if not os.path.exists(analyze_script):
            print(f"ERROR: analyze_offer2_template.py not found at {analyze_script}", flush=True)
            return False
        
        result = subprocess.run(
            ['python', analyze_script],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=120
        )
        
        if result.stdout:
            print(result.stdout, flush=True)
        if result.stderr:
            print(result.stderr, flush=True)
        
        if result.returncode != 0:
            print("ERROR: PROMPT 2 failed", flush=True)
            return False
        
        # Check if template_structure.json was created
        if os.path.exists(TEMPLATE_STRUCTURE_PATH):
            print("✓ PROMPT 2 completed successfully", flush=True)
            return True
        else:
            print("ERROR: PROMPT 2 did not create template_structure.json", flush=True)
            return False
        
    except subprocess.TimeoutExpired:
        print("ERROR: PROMPT 2 timed out", flush=True)
        return False
    except Exception as e:
        print(f"ERROR running PROMPT 2: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return False

def get_recomposition_plan(extraction_data, template_structure):
    """Get intelligent recomposition plan from GPT"""
    try:
        print("=" * 60, flush=True)
        print("GETTING RECOMPOSITION PLAN (REVISED PROMPT 3)", flush=True)
        print("=" * 60, flush=True)
        
        # Simplify data to avoid token limits
        simplified_extraction = {
            "items_count": len(extraction_data.get('items', [])),
            "items_preview": extraction_data.get('items', [])[:3],  # First 3 items as examples
            "technical_sections_count": len(extraction_data.get('technical_sections', [])),
            "technical_sections_preview": extraction_data.get('technical_sections', [])[:2],
            "has_images": len(extraction_data.get('images', [])) > 0,
            "document_metadata": extraction_data.get('document_metadata', {})
        }
        
        prompt = RECOMPOSITION_PROMPT.format(
            extraction_json=json.dumps(simplified_extraction, ensure_ascii=False, indent=2),
            template_structure_json=json.dumps(template_structure, ensure_ascii=False, indent=2)
        )
        
        print(f"Prompt length: {len(prompt)} characters", flush=True)
        print("Calling GPT-4o for recomposition planning...", flush=True)
        
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=3000,
            temperature=0.1  # Low temperature for consistent structure
        )
        
        plan_json = response.choices[0].message.content.strip()
        
        # Aggressive JSON cleaning
        if plan_json.startswith("```json"):
            plan_json = plan_json.replace("```json", "").replace("```", "").strip()
        elif plan_json.startswith("```"):
            plan_json = plan_json.replace("```", "").strip()
        
        # Remove any leading/trailing whitespace and control characters
        plan_json = plan_json.strip()
        
        # Log the cleaned response
        print(f"Raw response (first 500 chars):\n{plan_json[:500]}\n", flush=True)
        
        # Validate it looks like JSON
        if not plan_json.startswith('{'):
            print(f"ERROR: Response doesn't start with '{{'. Starts with: '{plan_json[:50]}'", flush=True)
            print(f"Full response:\n{plan_json}\n", flush=True)
            return None
        
        # Parse JSON
        try:
            recomposition_plan = json.loads(plan_json)
        except json.JSONDecodeError as e:
            print(f"JSON decode error: {e}", flush=True)
            print(f"Full response:\n{plan_json}", flush=True)
            return None
        
        # Validate required fields
        if not isinstance(recomposition_plan, dict):
            print(f"ERROR: Response is not a dictionary, it's {type(recomposition_plan)}", flush=True)
            print(f"Content: {recomposition_plan}", flush=True)
            return None
        
        # Check for required keys
        required_keys = ['document_sections']
        missing_keys = [key for key in required_keys if key not in recomposition_plan]
        
        if missing_keys:
            print(f"WARNING: Missing keys in response: {missing_keys}", flush=True)
            print(f"Available keys: {list(recomposition_plan.keys())}", flush=True)
            # Continue anyway - we can work with partial data
        
        print("✓ Recomposition plan created", flush=True)
        
        # Safe access to fields
        product_summary = recomposition_plan.get('product_summary', 'N/A')
        if product_summary and len(product_summary) > 60:
            print(f"  Product: {product_summary[:60]}...", flush=True)
        else:
            print(f"  Product: {product_summary}", flush=True)
        
        sections = recomposition_plan.get('document_sections', [])
        print(f"  Sections to create: {len(sections)}", flush=True)
        
        if sections:
            print("  Section types:", flush=True)
            for section in sections:
                print(f"    - {section.get('section_type', 'unknown')}: {section.get('position', 'unknown')}", flush=True)
        
        print("=" * 60, flush=True)
        
        return recomposition_plan
        
    except json.JSONDecodeError as e:
        print(f"JSON Parse Error: {e}", flush=True)
        print(f"Response was: {plan_json if 'plan_json' in locals() else 'No response'}", flush=True)
        return None
    except KeyError as e:
        print(f"KeyError accessing field: {e}", flush=True)
        print(f"Available keys: {list(recomposition_plan.keys()) if 'recomposition_plan' in locals() else 'N/A'}", flush=True)
        return None
    except Exception as e:
        print(f"Error getting recomposition plan: {e}", flush=True)
        print(f"Error type: {type(e).__name__}", flush=True)
        import traceback
        traceback.print_exc()
        return None

def execute_recomposition_docx(template_path, extraction_data, recomposition_plan, output_path):
    """Execute recomposition - REMOVE ALL Offer 2 equipment content, INSERT ALL Offer 1 content"""
    try:
        print("=" * 60, flush=True)
        print("EXECUTING DOCUMENT RECOMPOSITION", flush=True)
        print("=" * 60, flush=True)
        
        doc = Document(template_path)
        
        # STEP 1: IDENTIFY PRICING TABLE AND REMOVE ALL OTHER TABLES
        print("STEP 1: Identifying pricing table and removing Offer 2 tables...", flush=True)
        
        pricing_table = None
        tables_to_remove = []
        
        for table in doc.tables:
            # Pricing table has: multiple columns (>=4), multiple rows
            if len(table.columns) >= 4 and len(table.rows) >= 2:
                if pricing_table is None:
                    pricing_table = table
                    print(f"  ✓ Found pricing table: {len(table.rows)} rows, {len(table.columns)} cols", flush=True)
                else:
                    tables_to_remove.append(table)
            else:
                tables_to_remove.append(table)
        
        if not pricing_table:
            print("ERROR: No pricing table found", flush=True)
            return False
        
        # Remove all other tables (these are Offer 2 technical tables)
        for table in tables_to_remove:
            table._element.getparent().remove(table._element)
        
        print(f"  ✓ Removed {len(tables_to_remove)} extra tables from Offer 2", flush=True)
        
        # STEP 2: CLEAR PRICING TABLE COMPLETELY (including header with Offer 2 content)
        print("STEP 2: Clearing pricing table from Offer 2 content...", flush=True)
        
        original_cols = len(pricing_table.columns)
        print(f"  Original table has {original_cols} columns", flush=True)
        
        # Clear ALL rows (including header row which has "WORKING RANGE FOR CHAMPAGNE")
        while len(pricing_table.rows) > 0:
            pricing_table._tbl.remove(pricing_table.rows[0]._tr)
        
        print(f"  ✓ Cleared all rows from pricing table", flush=True)
        
        # Create new header row with generic structure
        header_row = pricing_table.add_row().cells
        if len(header_row) >= 5:
            header_row[0].text = "Pos"
            header_row[1].text = "Description"
            header_row[2].text = "Unit Price"
            header_row[3].text = "Qty"
            header_row[4].text = "Total"
            
            # Make header bold
            for cell in header_row[:5]:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.bold = True
                        run.font.size = Pt(11)
        
        print(f"  ✓ Created new clean header", flush=True)
        
        # STEP 3: REMOVE ALL OFFER 2 PARAGRAPHS (except company header/logo)
        print("STEP 3: Removing all Offer 2 paragraphs...", flush=True)
        
        # Keep only first 2 paragraphs (company header/logo)
        paragraphs_to_keep = 2
        removed_count = 0
        
        # Get all paragraph elements
        all_elements = list(doc.element.body)
        
        for i in range(len(all_elements) - 1, -1, -1):
            element = all_elements[i]
            if element.tag.endswith('p'):
                # Count from beginning
                para_position = sum(1 for e in all_elements[:i+1] if e.tag.endswith('p'))
                
                # Keep only first N paragraphs
                if para_position > paragraphs_to_keep:
                    element.getparent().remove(element)
                    removed_count += 1
        
        print(f"  ✓ Removed {removed_count} paragraphs (kept {paragraphs_to_keep} for header)", flush=True)
        
        # STEP 4: INSERT ALL OFFER 1 TECHNICAL CONTENT
        print("STEP 4: Inserting ALL Offer 1 technical content...", flush=True)
        
        technical_sections = extraction_data.get('technical_sections', [])
        
        if technical_sections and len(technical_sections) > 0:
            # Re-find table after cleanup
            pricing_table = doc.tables[0] if doc.tables else None
            if not pricing_table:
                print("ERROR: Lost pricing table during cleanup", flush=True)
                return False
            
            table_element = pricing_table._element
            table_parent = table_element.getparent()
            table_index = list(table_parent).index(table_element)
            
            print(f"  Inserting {len(technical_sections)} technical sections from Offer 1...", flush=True)
            
            for section in technical_sections:
                title = section.get('section_title', '') or section.get('title', '')
                content = section.get('content', '')
                content_type = section.get('type', 'text_paragraph')
                
                if not content or len(str(content).strip()) == 0:
                    continue
                
                # Add spacing
                space = doc.add_paragraph()
                table_parent.insert(table_index, space._element)
                table_index += 1
                
                # Add title
                if title:
                    title_para = doc.add_paragraph()
                    run = title_para.add_run(title)
                    run.font.bold = True
                    run.font.size = Pt(12)
                    table_parent.insert(table_index, title_para._element)
                    table_index += 1
                
                # Add content based on type
                if content_type == 'specification_table':
                    # This is a technical data table - insert as actual table
                    print(f"    Inserting specification table: {title}", flush=True)
                    
                    if isinstance(content, dict):
                        headers = content.get('table_headers', [])
                        rows = content.get('table_rows', [])
                        
                        if headers and rows:
                            # Create table
                            spec_table = doc.add_table(rows=1 + len(rows), cols=len(headers))
                            
                            # Try to apply a style, but don't fail if it doesn't exist
                            try:
                                spec_table.style = 'Light Grid Accent 1'
                            except:
                                try:
                                    spec_table.style = 'Table Grid'
                                except:
                                    pass  # Use default style
                            
                            # Headers
                            for i, header in enumerate(headers):
                                cell = spec_table.rows[0].cells[i]
                                cell.text = str(header)
                                for para in cell.paragraphs:
                                    for run in para.runs:
                                        run.font.bold = True
                                        run.font.size = Pt(10)
                            
                            # Data rows
                            for row_idx, row_data in enumerate(rows):
                                for col_idx, cell_data in enumerate(row_data):
                                    cell = spec_table.rows[row_idx + 1].cells[col_idx]
                                    cell.text = str(cell_data)
                                    # Make text smaller
                                    for para in cell.paragraphs:
                                        for run in para.runs:
                                            run.font.size = Pt(9)
                            
                            table_parent.insert(table_index, spec_table._element)
                            table_index += 1
                
                elif content_type == 'bullet_list' or isinstance(content, list):
                    # Bullet list
                    lines = content if isinstance(content, list) else str(content).split('\n')
                    for line in lines:
                        line = str(line).strip()
                        if line:
                            para = doc.add_paragraph(line, style='List Bullet')
                            table_parent.insert(table_index, para._element)
                            table_index += 1
                
                else:
                    # Regular text paragraph
                    para = doc.add_paragraph(str(content).strip())
                    for run in para.runs:
                        run.font.size = Pt(10)
                    table_parent.insert(table_index, para._element)
                    table_index += 1
            
            print(f"  ✓ Inserted all technical sections from Offer 1", flush=True)
        else:
            print("  ⚠ No technical sections found in Offer 1", flush=True)
        
        # STEP 5: FILL PRICING TABLE WITH OFFER 1 ITEMS
        print("STEP 5: Filling pricing table with Offer 1 items...", flush=True)
        
        items = extraction_data.get('items', [])
        
        if not items or len(items) == 0:
            print("  ⚠ WARNING: No items found in extraction!", flush=True)
            print("  This means PROMPT 1 failed to extract items properly", flush=True)
            return False
        
        print(f"  Processing {len(items)} items from Offer 1", flush=True)
        
        # Group by category
        categorized = OrderedDict()
        for item in items:
            cat = item.get('category', 'Items')
            if cat not in categorized:
                categorized[cat] = []
            categorized[cat].append(item)
        
        item_counter = 1
        for category, cat_items in categorized.items():
            # Category header
            cat_row = pricing_table.add_row().cells
            if len(cat_row) >= 2:
                cat_row[1].text = category
                for para in cat_row[1].paragraphs:
                    for run in para.runs:
                        run.font.bold = True
                        run.font.size = Pt(11)
            
            # Items
            for item in cat_items:
                row = pricing_table.add_row().cells
                
                # Position
                if len(row) >= 1:
                    row[0].text = str(item_counter)
                
                # Full description
                if len(row) >= 2:
                    desc_parts = []
                    
                    if item.get('item_name'):
                        desc_parts.append(item['item_name'])
                    if item.get('technical_description'):
                        desc_parts.append(item['technical_description'])
                    if item.get('specifications'):
                        if isinstance(item['specifications'], dict):
                            specs = ", ".join([f"{k}: {v}" for k, v in item['specifications'].items()])
                            desc_parts.append(f"Specifications: {specs}")
                        else:
                            desc_parts.append(str(item['specifications']))
                    if item.get('notes'):
                        desc_parts.append(item['notes'])
                    if item.get('details'):
                        desc_parts.append(item['details'])
                    
                    row[1].text = "\n\n".join(desc_parts)
                    for para in row[1].paragraphs:
                        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                # Prices
                if len(row) >= 3:
                    row[2].text = str(item.get('unit_price', ''))
                if len(row) >= 4:
                    row[3].text = str(item.get('quantity', '1'))
                if len(row) >= 5:
                    row[4].text = str(item.get('total_price', ''))
                
                item_counter += 1
        
        print(f"  ✓ Added {item_counter - 1} items to pricing table", flush=True)
        
        # STEP 6: SAVE
        doc.save(output_path)
        file_size = os.path.getsize(output_path)
        
        print("=" * 60, flush=True)
        print(f"✅ DOCUMENT SAVED: {output_path}", flush=True)
        print(f"   File size: {file_size:,} bytes", flush=True)
        print(f"   Items: {item_counter - 1}", flush=True)
        print(f"   Tables: {len(doc.tables)} (should be 1 pricing + N technical)", flush=True)
        print(f"   Paragraphs: {len(doc.paragraphs)}", flush=True)
        print("=" * 60, flush=True)
        
        return True
        
    except Exception as e:
        print(f"ERROR: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return False

def generate_offer_flexible():
    """Main generation function with auto PROMPT 2 execution"""
    try:
        print("=== FLEXIBLE GENERATION (RECOMPOSITION SYSTEM) ===", flush=True)
        
        if not openai.api_key:
            print("ERROR: OPENAI_API_KEY not set", flush=True)
            return False
        
        # STEP 1: Load extraction data
        extraction_data = load_extraction_data()
        if not extraction_data:
            print("ERROR: No extraction data", flush=True)
            return False
        
        print(f"✓ Loaded extraction: {len(extraction_data.get('items', []))} items", flush=True)
        
        # STEP 2: Load template structure - AUTO-RUN PROMPT 2 IF MISSING
        template_structure = load_template_structure()
        
        if not template_structure:
            print("⚠ Template structure missing, running PROMPT 2 now...", flush=True)
            
            # Find template file
            template_docx = os.path.join(BASE_DIR, "offer2_template.docx")
            template_xlsx = os.path.join(BASE_DIR, "offer2_template.xlsx")
            
            if not os.path.exists(template_docx) and not os.path.exists(template_xlsx):
                print("ERROR: No template found (need offer2_template.docx or .xlsx)", flush=True)
                return False
            
            # Run PROMPT 2
            success = run_prompt_2_analysis()
            
            if not success:
                print("ERROR: PROMPT 2 failed to analyze template", flush=True)
                return False
            
            # Try loading template structure again
            template_structure = load_template_structure()
            
            if not template_structure:
                print("ERROR: PROMPT 2 completed but template_structure.json is invalid", flush=True)
                return False
        
        print("✓ Template structure loaded", flush=True)
        
        # STEP 3: Get recomposition plan from GPT
        recomposition_plan = get_recomposition_plan(extraction_data, template_structure)
        
        if not recomposition_plan:
            print("⚠ GPT recomposition plan failed, using simple default plan", flush=True)
            # Use simple default plan
            recomposition_plan = {
                "product_summary": "Equipment quotation",
                "document_sections": [
                    {
                        "section_type": "pricing_table",
                        "position": "main_body",
                        "content": {
                            "categories": ["Items"],
                            "column_mapping": {
                                "col_0": "position_number",
                                "col_1": "full_description_with_specs",
                                "col_2": "unit_price",
                                "col_3": "quantity",
                                "col_4": "total_price"
                            },
                            "items_to_include": "all"
                        }
                    }
                ],
                "content_mapping": {
                    "technical_descriptions_placement": "in_table"
                },
                "formatting_instructions": {},
                "translation_needed": False
            }
            print("✓ Using fallback plan", flush=True)
        
        # Save plan for debugging
        plan_path = os.path.join(OUTPUT_FOLDER, "recomposition_plan.json")
        with open(plan_path, 'w', encoding='utf-8') as f:
            json.dump(recomposition_plan, f, indent=2, ensure_ascii=False)
        print(f"✓ Saved recomposition plan: {plan_path}", flush=True)
        
        # STEP 4: Find template file
        template_docx = os.path.join(BASE_DIR, "offer2_template.docx")
        template_xlsx = os.path.join(BASE_DIR, "offer2_template.xlsx")
        
        if os.path.exists(template_docx):
            template_path = template_docx
        elif os.path.exists(template_xlsx):
            print("ERROR: XLSX templates not yet supported in recomposition system", flush=True)
            print("Please use DOCX template", flush=True)
            return False
        else:
            print("ERROR: No template file found", flush=True)
            return False
        
        # STEP 5: Execute recomposition
        output_path = os.path.join(OUTPUT_FOLDER, "final_offer1.docx")
        success = execute_recomposition_docx(template_path, extraction_data, recomposition_plan, output_path)
        
        if success:
            print("=== RECOMPOSITION COMPLETED SUCCESSFULLY ===", flush=True)
            return True
        else:
            print("ERROR: Recomposition execution failed", flush=True)
            return False
        
    except Exception as e:
        print(f"FATAL ERROR: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Recomposition Generation Script Started (REVISED PROMPT 3)", flush=True)
    
    success = generate_offer_flexible()
    
    if not success:
        print("Generation failed", flush=True)
        sys.exit(1)
    
    print("COMPLETED SUCCESSFULLY", flush=True)
    sys.exit(0)