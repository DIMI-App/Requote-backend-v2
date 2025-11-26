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
        
        # Clean JSON
        if plan_json.startswith("```json"):
            plan_json = plan_json.replace("```json", "").replace("```", "").strip()
        elif plan_json.startswith("```"):
            plan_json = plan_json.replace("```", "").strip()
        
        print(f"Raw response (first 500 chars):\n{plan_json[:500]}\n", flush=True)
        
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
    """Execute the recomposition plan to create final DOCX"""
    try:
        print("=" * 60, flush=True)
        print("EXECUTING RECOMPOSITION PLAN", flush=True)
        print("=" * 60, flush=True)
        
        doc = Document(template_path)
        
        # Find pricing table (still need to fill it)
        pricing_table = None
        for table in doc.tables:
            if len(table.rows) > 1 and len(table.columns) >= 3:
                pricing_table = table
                break
        
        if not pricing_table:
            print("WARNING: No pricing table found, creating basic document", flush=True)
            # Would need to create document from scratch - complex, skip for now
            return False
        
        print(f"✓ Found pricing table: {len(pricing_table.rows)} rows", flush=True)
        
        # Clear pricing table (keep header)
        while len(pricing_table.rows) > 1:
            pricing_table._tbl.remove(pricing_table.rows[1]._tr)
        
        print("✓ Cleared existing table data", flush=True)
        
        # Get column mapping from plan
        column_mapping = {}
        for section in recomposition_plan.get('document_sections', []):
            if section.get('section_type') == 'pricing_table':
                column_mapping = section.get('content', {}).get('column_mapping', {})
                break
        
        print(f"✓ Column mapping: {column_mapping}", flush=True)
        
        # Fill table with items
        items = extraction_data.get('items', [])
        print(f"✓ Filling table with {len(items)} items", flush=True)
        
        # Group by category
        categorized = OrderedDict()
        for item in items:
            cat = item.get('category', 'Items')
            if cat not in categorized:
                categorized[cat] = []
            categorized[cat].append(item)
        
        print(f"✓ Grouped into {len(categorized)} categories", flush=True)
        
        item_counter = 1
        for category, cat_items in categorized.items():
            print(f"  Processing category: {category} ({len(cat_items)} items)", flush=True)
            
            # Category row
            cat_row = pricing_table.add_row().cells
            if len(cat_row) >= 2:
                cat_row[1].text = category
                for para in cat_row[1].paragraphs:
                    for run in para.runs:
                        run.font.bold = True
                        run.font.size = Pt(11)
            
            # Item rows
            for item in cat_items:
                row = pricing_table.add_row().cells
                
                # Map according to plan (or use default)
                if len(row) >= 1:
                    row[0].text = str(item_counter)
                
                if len(row) >= 2:
                    # Build FULL description from ALL available fields
                    desc_parts = []
                    
                    # Item name
                    if item.get('item_name'):
                        desc_parts.append(item['item_name'])
                    
                    # Technical description
                    if item.get('technical_description'):
                        desc_parts.append(item['technical_description'])
                    
                    # Specifications
                    if item.get('specifications'):
                        if isinstance(item['specifications'], dict):
                            specs = ", ".join([f"{k}: {v}" for k, v in item['specifications'].items()])
                            desc_parts.append(f"\nSpecifications: {specs}")
                        elif isinstance(item['specifications'], str):
                            desc_parts.append(f"\n{item['specifications']}")
                    
                    # Notes
                    if item.get('notes'):
                        desc_parts.append(f"\n{item['notes']}")
                    
                    # Legacy details field
                    if item.get('details'):
                        desc_parts.append(f"\n{item['details']}")
                    
                    full_description = "\n\n".join(desc_parts)
                    row[1].text = full_description
                
                if len(row) >= 3:
                    row[2].text = str(item.get('unit_price', ''))
                
                if len(row) >= 4:
                    row[3].text = str(item.get('quantity', '1'))
                
                if len(row) >= 5:
                    row[4].text = str(item.get('total_price', ''))
                
                item_counter += 1
        
        print(f"✓ Filled table with all items", flush=True)
        
        # Save
        doc.save(output_path)
        file_size = os.path.getsize(output_path)
        
        print(f"✓ Document saved: {output_path}", flush=True)
        print(f"  File size: {file_size:,} bytes", flush=True)
        print("=" * 60, flush=True)
        
        return True
        
    except Exception as e:
        print(f"ERROR executing recomposition: {e}", flush=True)
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
            print("ERROR: Could not create recomposition plan", flush=True)
            return False
        
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