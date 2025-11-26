import os
import sys
import json
import openai
import fitz
import base64
from docx import Document
import openpyxl

openai.api_key = os.environ.get('OPENAI_API_KEY')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# PROMPT 2: Analyze Offer 2 Template Structure
TEMPLATE_ANALYSIS_PROMPT = """PROMPT 2: ANALYZE OFFER 2 TEMPLATE STRUCTURE
==============================================

You are a document formatting expert who analyzes company quotation templates to understand their structure, so you can replicate the same format when inserting new data.

YOUR TASK:
A company has given you their branded quotation template (Offer 2). You need to thoroughly analyze it to understand:
1. How it's structured
2. Where different content types are placed
3. What formatting style is used
4. How to preserve this exact look when filling it with new data

ANALYZE THE FOLLOWING ASPECTS:

1. DOCUMENT STRUCTURE & FLOW
   - What comes first? (Company header, logo, client info, intro text?)
   - What's the sequence? (Description → Price table → Photos → Terms?)
   - How many distinct sections are there?
   - Where does pricing table appear relative to descriptions?
   - Where are images placed? (Before table, after table, mixed in?)

2. HEADER & BRANDING
   - Company logo position (top-left, center, top-right?)
   - Company name, address, contact information placement
   - Header style (simple text, formatted box, letterhead?)
   - Client information section (where and how formatted?)
   - Date, quotation number, reference fields

3. PRICING TABLE STRUCTURE
   - Column headers (in what language?)
   - Column order (Item, Qty, Price, Total, Notes?)
   - How many columns exactly?
   - Is there a "Category" or "Section" column?
   - Is there a "Description" or "Technical specs" column?
   - Where are subtotals? (After each category, only at end?)
   - Grand total placement and formatting
   - Currency symbol position

4. TEXT CONTENT PLACEMENT
   - Is there introductory text before the price table?
   - Are there product descriptions above/below table?
   - Are technical specifications in separate sections?
   - How are features/benefits presented? (Paragraphs, bullets, tables?)
   - Is there explanatory text between pricing categories?

5. VISUAL ELEMENTS
   - Are there product images? Where exactly?
   - Are images next to items or in separate section?
   - Image size and alignment (left, right, center, inline?)
   - Are there logos, watermarks, or background graphics?
   - Charts, diagrams, technical drawings?

6. FORMATTING STYLE
   - Font family and sizes for different elements
   - Bold, italic, underline usage patterns
   - Color scheme (headers, table, text, borders)
   - Line spacing and paragraph spacing
   - Borders and shading (table cells, sections)
   - Page margins and layout

7. TABLE FORMATTING DETAILS
   - Header row style (background color, bold text, borders)
   - Data row style (alternating colors, borders, alignment)
   - Text alignment in each column (left, right, center)
   - Number formatting (thousand separators, decimals)
   - Currency format (€1.234,56 or €1,234.56 or € 1,234.56)

8. SECTIONS & CATEGORIES
   - How are different product categories separated?
   - Are there section headers within the table?
   - Separate tables for Main/Optional/Accessories?
   - Or one unified table with category labels?

9. FOOTER & CLOSING
   - Terms and conditions (where placed?)
   - Payment terms, delivery information
   - Signatures, approval sections
   - Contact information repeated at bottom?

10. MULTI-PAGE BEHAVIOR
    - How does content flow across pages?
    - Are headers/footers repeated on each page?
    - Page numbers and their position
    - Does table break across pages or stay on one?

CRITICAL: Return a comprehensive JSON structure that captures ALL these details so that another system can perfectly recreate this template format when filling it with new data.

NOW ANALYZE THIS TEMPLATE:"""

def analyze_docx_template(template_path):
    """Analyze DOCX template using PROMPT 2"""
    try:
        print("Analyzing DOCX template...", flush=True)
        doc = Document(template_path)
        
        analysis_text = f"\nDOCUMENT TYPE: DOCX\n"
        analysis_text += f"PARAGRAPHS: {len(doc.paragraphs)}\n"
        analysis_text += f"TABLES: {len(doc.tables)}\n\n"
        
        # Get sample text
        analysis_text += "SAMPLE CONTENT:\n"
        for i, para in enumerate(doc.paragraphs[:10]):
            if para.text.strip():
                analysis_text += f"Para {i+1}: {para.text[:100]}\n"
        
        # Get table structure
        if doc.tables:
            analysis_text += f"\nTABLE STRUCTURE (First table):\n"
            table = doc.tables[0]
            analysis_text += f"Rows: {len(table.rows)}, Columns: {len(table.columns)}\n"
            if table.rows:
                analysis_text += "Header row: " + " | ".join([cell.text[:20] for cell in table.rows[0].cells]) + "\n"
        
        return analysis_text
        
    except Exception as e:
        print(f"Error analyzing DOCX: {e}", flush=True)
        return f"DOCX template structure: {str(e)}"

def analyze_xlsx_template(template_path):
    """Analyze XLSX template using PROMPT 2"""
    try:
        print("Analyzing XLSX template...", flush=True)
        workbook = openpyxl.load_workbook(template_path)
        sheet = workbook.active
        
        analysis_text = f"\nDOCUMENT TYPE: XLSX\n"
        analysis_text += f"SHEET NAME: {sheet.title}\n"
        analysis_text += f"MAX ROW: {sheet.max_row}, MAX COL: {sheet.max_column}\n\n"
        
        # Get sample content
        analysis_text += "SAMPLE CONTENT (First 10 rows):\n"
        for row_idx in range(1, min(11, sheet.max_row + 1)):
            row = sheet[row_idx]
            row_values = [str(cell.value) if cell.value else '' for cell in row]
            analysis_text += f"Row {row_idx}: " + " | ".join(row_values[:5]) + "\n"
        
        return analysis_text
        
    except Exception as e:
        print(f"Error analyzing XLSX: {e}", flush=True)
        return f"XLSX template structure: {str(e)}"

def create_fallback_structure(template_path, output_path):
    """Create a basic fallback structure when GPT analysis fails"""
    try:
        file_ext = template_path.lower().split('.')[-1]
        
        print("Creating fallback template structure (basic)...", flush=True)
        
        fallback_structure = {
            "analysis_method": "FALLBACK_Basic_Structure",
            "template_file": template_path,
            "template_type": file_ext,
            "structure": {
                "document_type": file_ext.upper(),
                "template_language": "EN",
                "pricing_table": {
                    "location": "main_table",
                    "columns": [
                        {"name": "Position", "type": "number", "alignment": "center"},
                        {"name": "Description", "type": "text", "alignment": "left"},
                        {"name": "Unit Price", "type": "currency", "alignment": "right"},
                        {"name": "Quantity", "type": "number", "alignment": "center"},
                        {"name": "Total Price", "type": "currency", "alignment": "right"}
                    ],
                    "currency_format": "€1,234.56"
                },
                "content_placement": {
                    "descriptions": "in_table_column",
                    "images": "not_applicable",
                    "technical_specs": "in_description_column"
                },
                "formatting_rules": {
                    "category_style": "bold_row",
                    "header_style": "bold_colored",
                    "number_format": "standard"
                }
            }
        }
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(fallback_structure, f, indent=2, ensure_ascii=False)
        
        print(f"✓ Fallback structure saved to {output_path}", flush=True)
        return True
        
    except Exception as e:
        print(f"ERROR creating fallback: {e}", flush=True)
        return False

def analyze_template_structure(template_path, output_path):
    """Main function to analyze template using PROMPT 2 and GPT-4o"""
    try:
        print("=== STARTING TEMPLATE ANALYSIS (PROMPT 2) ===", flush=True)
        
        if not openai.api_key:
            print("ERROR: OPENAI_API_KEY not set", flush=True)
            # Create fallback structure
            create_fallback_structure(template_path, output_path)
            return False
        
        if not os.path.exists(template_path):
            print(f"ERROR: Template not found: {template_path}", flush=True)
            return False
        
        # Determine template type
        file_ext = template_path.lower().split('.')[-1]
        print(f"Template type: {file_ext}", flush=True)
        
        # Get structural information
        if file_ext in ['docx', 'doc']:
            structure_info = analyze_docx_template(template_path)
        elif file_ext in ['xlsx', 'xls']:
            structure_info = analyze_xlsx_template(template_path)
        else:
            print(f"Unsupported template type: {file_ext}", flush=True)
            create_fallback_structure(template_path, output_path)
            return False
        
        # Build prompt with structure info
        full_prompt = TEMPLATE_ANALYSIS_PROMPT + "\n\n" + structure_info
        
        print("Calling GPT-4o for template analysis...", flush=True)
        
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": full_prompt}],
            max_tokens=4000,
            temperature=0
        )
        
        print("Received template analysis", flush=True)
        
        analysis_json = response.choices[0].message.content.strip()
        
        # Clean JSON formatting
        if analysis_json.startswith("```json"):
            analysis_json = analysis_json.replace("```json", "").replace("```", "").strip()
        elif analysis_json.startswith("```"):
            analysis_json = analysis_json.replace("```", "").strip()
        
        print("Parsing template analysis JSON...", flush=True)
        template_structure = json.loads(analysis_json)
        
        # Save analysis
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        output_data = {
            "analysis_method": "PROMPT_2_Template_Analysis",
            "template_file": template_path,
            "template_type": file_ext,
            "structure": template_structure
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False)
        
        print(f"✓ Template analysis saved to {output_path}", flush=True)
        print("=== TEMPLATE ANALYSIS COMPLETED ===", flush=True)
        
        # Print summary
        if isinstance(template_structure, dict):
            print("\nTemplate Summary:", flush=True)
            if 'document_type' in template_structure:
                print(f"  Type: {template_structure['document_type']}", flush=True)
            if 'template_language' in template_structure:
                print(f"  Language: {template_structure['template_language']}", flush=True)
            if 'pricing_table' in template_structure and isinstance(template_structure['pricing_table'], dict):
                if 'columns' in template_structure['pricing_table']:
                    print(f"  Columns: {len(template_structure['pricing_table']['columns'])}", flush=True)
        
        return True
        
    except Exception as e:
        print(f"ERROR in template analysis: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        
        # Create fallback structure so generation doesn't fail
        print("Creating fallback template structure due to error...", flush=True)
        create_fallback_structure(template_path, output_path)
        return False

if __name__ == "__main__":
    print("Template Analysis Script Started (PROMPT 2)", flush=True)
    
    # Try to find template (check both DOCX and XLSX)
    template_docx = os.path.join(BASE_DIR, "offer2_template.docx")
    template_xlsx = os.path.join(BASE_DIR, "offer2_template.xlsx")
    
    if os.path.exists(template_docx):
        template_path = template_docx
    elif os.path.exists(template_xlsx):
        template_path = template_xlsx
    else:
        print("ERROR: No template found (offer2_template.docx or .xlsx)", flush=True)
        sys.exit(1)
    
    output_path = os.path.join(OUTPUT_FOLDER, "template_structure.json")
    
    success = analyze_template_structure(template_path, output_path)
    
    # IMPORTANT: Always exit 0 if we created ANY structure (even fallback)
    # This prevents the generation step from failing completely
    if os.path.exists(output_path):
        print("COMPLETED (template_structure.json exists)", flush=True)
        sys.exit(0)
    else:
        print("FAILED (no template_structure.json created)", flush=True)
        sys.exit(1)