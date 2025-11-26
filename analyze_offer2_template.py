import os
import sys
import json
import openai
import fitz
from docx import Document
import openpyxl

openai.api_key = os.environ.get('OPENAI_API_KEY')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# REVISED PROMPT 2: DEEP TEMPLATE STRUCTURE ANALYSIS
TEMPLATE_ANALYSIS_PROMPT = """You are analyzing a company's quotation template to understand its COMPLETE structure, not just the pricing table.

YOUR TASK: Understand this template like a human sales manager who needs to fill it out manually.

ANALYZE THESE ASPECTS:

1. DOCUMENT FLOW & SECTIONS
   - What comes first? (Company header, logo, client address section?)
   - What's the sequence? (Header → Intro text → Technical specs → Pricing table → Optional items → Terms?)
   - How many distinct sections are there?
   - What is the narrative flow?

2. COMPANY BRANDING
   - Company name and location
   - Logo position (if visible)
   - Header style (letterhead, simple text, formatted box?)
   - Contact information placement
   - Brand colors or styling

3. CONTENT PLACEMENT STRATEGY
   - WHERE do product descriptions go?
     * Before pricing table as paragraphs?
     * Inside table description column?
     * After table as appendix?
     * Separate technical section?
   
   - WHERE do technical specifications go?
     * Integrated in descriptions?
     * Separate specifications table?
     * Bullet points before/after main table?
     * Technical data sheet at end?
   
   - WHERE do images/photos go?
     * Inline with descriptions?
     * Separate section at end?
     * Next to each item?
     * Not included?
   
   - WHERE do feature lists go?
     * Bullet points in description?
     * Separate "Features" section?
     * Integrated in table?

4. PRICING TABLE STRUCTURE
   - Table location in document (after what? before what?)
   - Column headers (exact text, in what language)
   - Column count and purposes
   - How are categories shown? (header rows? separate sections?)
   - How are optional items shown? (separate table? marked in main table?)
   - Currency format (€1.234,56 or $1,234.56)
   - Number formatting

5. TEXT SECTIONS
   - Is there introductory text before table?
   - Are there explanatory paragraphs between sections?
   - Is there a conclusion/summary after table?
   - Payment terms location?
   - Delivery terms location?

6. SPECIAL SECTIONS
   - "Included" items section?
   - "Optional" or "Accessories" section?
   - "Exclusions" section?
   - Terms and conditions?
   - Warranty information?

7. FORMATTING PATTERNS
   - Font family used
   - How are section headers formatted? (bold, size, color)
   - How is body text formatted?
   - Line spacing and margins
   - Use of colors (headers, borders, backgrounds)

8. MULTI-LANGUAGE DETECTION
   - What language is the template in?
   - Are there any bilingual elements?

CRITICAL: Your analysis must be detailed enough that another person could recreate this template's EXACT structure and flow from your description.

Return comprehensive JSON that maps the complete document structure.

NOW ANALYZE THIS TEMPLATE:

Return JSON with this structure:
{
  "document_flow": {
    "sections_in_order": ["header", "intro_text", "technical_description", "pricing_table", "optional_items", "terms"],
    "narrative_style": "formal_technical / sales_oriented / minimal"
  },
  "company_branding": {
    "company_name": "...",
    "template_language": "EN/UK/IT/ES/DE/FR",
    "header_style": "description"
  },
  "content_placement": {
    "product_descriptions": "before_table / in_table_column / after_table / separate_section",
    "technical_specs": "in_description / separate_specs_table / bullet_points / technical_data_sheet",
    "images": "inline / after_table / not_included / separate_page",
    "feature_lists": "bullets_in_description / separate_features_section / in_table"
  },
  "pricing_table": {
    "location": "after_technical_description",
    "columns": [
      {"name": "POS.", "purpose": "position_number"},
      {"name": "DESCRIPTION", "purpose": "item_description"},
      {"name": "Q.", "purpose": "quantity"},
      {"name": "UNIT PRICE", "purpose": "unit_price"},
      {"name": "TOTAL", "purpose": "total_price"}
    ],
    "categories_shown_as": "bold_header_rows / separate_tables / column",
    "optional_items": "separate_section_after_main / marked_in_main_table / separate_table",
    "currency_format": "€1.234,56"
  },
  "text_sections": {
    "has_intro_before_table": true/false,
    "intro_content_type": "product_overview / company_intro / order_details",
    "has_text_between_sections": true/false,
    "has_conclusion_after_table": true/false,
    "payment_terms_location": "header_box / after_table / separate_section",
    "delivery_terms_location": "header_box / after_table / separate_section"
  },
  "special_sections": {
    "has_exclusions_section": true/false,
    "has_warranty_section": true/false,
    "has_technical_data_table": true/false,
    "technical_data_location": "before_pricing / after_pricing / separate_page"
  },
  "formatting_guide": {
    "primary_font": "Arial / Calibri / Times",
    "header_formatting": "bold_12pt / bold_colored_14pt",
    "body_text_formatting": "normal_10pt / normal_11pt",
    "emphasis_color": "#HEX if any"
  }
}
"""

def analyze_docx_template(template_path):
    """Deep analysis of DOCX template structure"""
    try:
        print("Analyzing DOCX template structure...", flush=True)
        doc = Document(template_path)
        
        analysis_text = f"\nTEMPLATE TYPE: DOCX\n"
        analysis_text += f"TOTAL PARAGRAPHS: {len(doc.paragraphs)}\n"
        analysis_text += f"TOTAL TABLES: {len(doc.tables)}\n\n"
        
        # Extract document flow
        analysis_text += "DOCUMENT STRUCTURE:\n"
        
        # First 20 paragraphs (to understand header and intro)
        analysis_text += "\nFIRST 20 PARAGRAPHS (Headers and intro):\n"
        for i, para in enumerate(doc.paragraphs[:20]):
            if para.text.strip():
                style = para.style.name if para.style else "Normal"
                analysis_text += f"  Para {i+1} [{style}]: {para.text[:100]}\n"
        
        # Analyze all tables
        analysis_text += f"\nTABLES ANALYSIS ({len(doc.tables)} total):\n"
        for table_idx, table in enumerate(doc.tables):
            analysis_text += f"\n  Table {table_idx + 1}:\n"
            analysis_text += f"    Rows: {len(table.rows)}, Columns: {len(table.columns)}\n"
            
            # Get header row
            if table.rows:
                headers = [cell.text[:30] for cell in table.rows[0].cells]
                analysis_text += f"    Headers: {' | '.join(headers)}\n"
            
            # Sample first 3 data rows
            if len(table.rows) > 1:
                analysis_text += f"    Sample rows:\n"
                for row_idx in range(1, min(4, len(table.rows))):
                    row_data = [cell.text[:20] for cell in table.rows[row_idx].cells]
                    analysis_text += f"      Row {row_idx}: {' | '.join(row_data)}\n"
        
        # Last 10 paragraphs (to understand footer/terms)
        analysis_text += "\nLAST 10 PARAGRAPHS (Footer and terms):\n"
        for i, para in enumerate(doc.paragraphs[-10:]):
            if para.text.strip():
                analysis_text += f"  Para {len(doc.paragraphs) - 10 + i + 1}: {para.text[:100]}\n"
        
        return analysis_text
        
    except Exception as e:
        print(f"Error analyzing DOCX: {e}", flush=True)
        return f"Error analyzing DOCX: {str(e)}"

def analyze_xlsx_template(template_path):
    """Deep analysis of XLSX template structure"""
    try:
        print("Analyzing XLSX template structure...", flush=True)
        workbook = openpyxl.load_workbook(template_path)
        sheet = workbook.active
        
        analysis_text = f"\nTEMPLATE TYPE: XLSX\n"
        analysis_text += f"SHEET: {sheet.title}\n"
        analysis_text += f"DIMENSIONS: {sheet.max_row} rows x {sheet.max_column} cols\n\n"
        
        # Get all non-empty rows
        analysis_text += "SHEET STRUCTURE:\n"
        for row_idx in range(1, min(30, sheet.max_row + 1)):
            row = sheet[row_idx]
            row_values = [str(cell.value) if cell.value else '' for cell in row]
            row_text = " | ".join([v[:30] for v in row_values if v.strip()])
            if row_text:
                analysis_text += f"  Row {row_idx}: {row_text}\n"
        
        return analysis_text
        
    except Exception as e:
        print(f"Error analyzing XLSX: {e}", flush=True)
        return f"Error analyzing XLSX: {str(e)}"

def analyze_template_structure(template_path, output_path):
    """Main analysis function"""
    try:
        print("=== STARTING DEEP TEMPLATE ANALYSIS (REVISED PROMPT 2) ===", flush=True)
        
        if not openai.api_key:
            print("ERROR: OPENAI_API_KEY not set", flush=True)
            return False
        
        if not os.path.exists(template_path):
            print(f"ERROR: Template not found: {template_path}", flush=True)
            return False
        
        # Determine file type
        file_ext = template_path.lower().split('.')[-1]
        print(f"Template type: {file_ext}", flush=True)
        
        # Get structural information
        if file_ext in ['docx', 'doc']:
            structure_info = analyze_docx_template(template_path)
        elif file_ext in ['xlsx', 'xls']:
            structure_info = analyze_xlsx_template(template_path)
        else:
            print(f"Unsupported template type: {file_ext}", flush=True)
            return False
        
        # Build full prompt
        full_prompt = TEMPLATE_ANALYSIS_PROMPT + "\n\n" + structure_info
        
        print("Calling GPT-4o for deep template analysis...", flush=True)
        print(f"Prompt length: {len(full_prompt)} characters", flush=True)
        
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": full_prompt}],
            max_tokens=4000,
            temperature=0
        )
        
        analysis_json = response.choices[0].message.content.strip()
        
        # Clean JSON
        if analysis_json.startswith("```json"):
            analysis_json = analysis_json.replace("```json", "").replace("```", "").strip()
        elif analysis_json.startswith("```"):
            analysis_json = analysis_json.replace("```", "").strip()
        
        print("Parsing template analysis...", flush=True)
        template_structure = json.loads(analysis_json)
        
        # Save
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        output_data = {
            "analysis_method": "REVISED_PROMPT_2_Deep_Structure",
            "template_file": template_path,
            "template_type": file_ext,
            "structure": template_structure
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False)
        
        print(f"✓ Deep template analysis saved: {output_path}", flush=True)
        print("=== TEMPLATE ANALYSIS COMPLETED ===", flush=True)
        
        # Print summary
        if isinstance(template_structure, dict):
            print("\nTemplate Structure Summary:", flush=True)
            if 'company_branding' in template_structure:
                print(f"  Company: {template_structure['company_branding'].get('company_name', 'Unknown')}", flush=True)
                print(f"  Language: {template_structure['company_branding'].get('template_language', 'Unknown')}", flush=True)
            if 'document_flow' in template_structure:
                sections = template_structure['document_flow'].get('sections_in_order', [])
                print(f"  Document Flow: {' → '.join(sections)}", flush=True)
            if 'content_placement' in template_structure:
                placement = template_structure['content_placement']
                print(f"  Descriptions: {placement.get('product_descriptions', 'unknown')}", flush=True)
                print(f"  Tech Specs: {placement.get('technical_specs', 'unknown')}", flush=True)
        
        return True
        
    except Exception as e:
        print(f"FATAL ERROR: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Revised Template Analysis Script Started (DEEP STRUCTURE)", flush=True)
    
    # Find template
    template_docx = os.path.join(BASE_DIR, "offer2_template.docx")
    template_xlsx = os.path.join(BASE_DIR, "offer2_template.xlsx")
    
    if os.path.exists(template_docx):
        template_path = template_docx
    elif os.path.exists(template_xlsx):
        template_path = template_xlsx
    else:
        print("ERROR: No template found", flush=True)
        sys.exit(1)
    
    output_path = os.path.join(OUTPUT_FOLDER, "template_structure.json")
    
    success = analyze_template_structure(template_path, output_path)
    
    if not success:
        print("Template analysis failed", flush=True)
        sys.exit(1)
    
    print("COMPLETED SUCCESSFULLY (REVISED PROMPT 2)", flush=True)
    sys.exit(0)