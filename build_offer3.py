"""
Build Offer 3 with STRUCTURE PRESERVATION
Reads content_blocks from JSON and rebuilds with tables, bullets, formatting
"""

import os
import sys
import json
import shutil
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from standard_template import Offer3Template

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')

def add_structured_content_to_doc(doc, items):
    """
    Add technical content with PRESERVED STRUCTURE from content_blocks
    """
    
    # Section heading
    heading = doc.add_paragraph()
    run = heading.add_run("Technical Specifications")
    run.font.size = Pt(14)
    run.font.bold = True
    
    doc.add_paragraph()  # Spacing
    
    items_added = 0
    
    for item in items:
        content_blocks = item.get('content_blocks', [])
        
        if not content_blocks:
            continue
        
        # Item heading
        item_heading = doc.add_paragraph()
        run = item_heading.add_run(f"{item['item_number']}. {item['item_name']}")
        run.font.size = Pt(12)
        run.font.bold = True
        
        # Rebuild content with structure
        for block in content_blocks:
            block_type = block.get('type', 'paragraph')
            
            if block_type == 'table':
                # Rebuild table
                table_data = block.get('data', [])
                
                if not table_data:
                    continue
                
                num_rows = len(table_data)
                num_cols = len(table_data[0]) if table_data else 0
                
                if num_rows > 0 and num_cols > 0:
                    table = doc.add_table(rows=num_rows, cols=num_cols)
                    table.style = 'Light Grid Accent 1'
                    
                    for row_idx, row_data in enumerate(table_data):
                        for col_idx, cell_text in enumerate(row_data):
                            if col_idx < num_cols:
                                cell_value = str(cell_text) if cell_text is not None else ''
                                table.rows[row_idx].cells[col_idx].text = cell_value
            
            elif block_type == 'bullet':
                # Bullet point
                para = doc.add_paragraph(style='List Bullet')
                run = para.add_run(block.get('text', ''))
                run.font.size = Pt(11)
            
            elif block_type == 'numbered_list':
                # Numbered list
                para = doc.add_paragraph(style='List Number')
                run = para.add_run(block.get('text', ''))
                run.font.size = Pt(11)
            
            elif block_type == 'heading':
                # Sub-heading
                para = doc.add_paragraph()
                run = para.add_run(block.get('text', ''))
                run.font.size = Pt(11)
                run.font.bold = True
            
            else:
                # Normal paragraph
                text = block.get('text', '')
                if text.strip():
                    para = doc.add_paragraph()
                    run = para.add_run(text)
                    run.font.size = Pt(11)
        
        doc.add_paragraph()  # Spacing between items
        items_added += 1
    
    print(f"  ✓ Added structured content for {items_added} items", flush=True)

def generate_offer3(company_data_path, items_data_path, output_path):
    """
    Build Offer 3 using structured content from extraction
    """
    
    try:
        print("=" * 60, flush=True)
        print("BUILDING OFFER 3 - STRUCTURE PRESERVED", flush=True)
        print("=" * 60, flush=True)
        
        # Load items data
        print(f"Loading items data: {items_data_path}", flush=True)
        if not os.path.exists(items_data_path):
            print("✗ Items data not found!", flush=True)
            return False
        
        with open(items_data_path, 'r', encoding='utf-8') as f:
            items_data = json.load(f)
        
        items = items_data.get('items', [])
        print(f"✓ Loaded {len(items)} items", flush=True)
        
        # Check structure preservation
        items_with_blocks = sum(1 for item in items if item.get('content_blocks'))
        total_blocks = sum(len(item.get('content_blocks', [])) for item in items)
        
        print(f"  Items with structured content: {items_with_blocks}/{len(items)}", flush=True)
        print(f"  Total content blocks: {total_blocks}", flush=True)
        
        # Load company data (optional)
        company_data = {}
        if os.path.exists(company_data_path):
            with open(company_data_path, 'r', encoding='utf-8') as f:
                company_data = json.load(f)
        
        # Find template file
        template_path_options = [
            os.path.join(BASE_DIR, 'offer2_template.docx'),
            os.path.join(BASE_DIR, 'uploads', 'offer2_template.docx'),
        ]
        
        template_path = None
        for path in template_path_options:
            if os.path.exists(path):
                template_path = path
                print(f"✓ Found template: {path}", flush=True)
                break
        
        if not template_path:
            print("✗ Template not found!", flush=True)
            return False
        
        # Copy template to output
        print(f"Copying template to: {output_path}", flush=True)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        shutil.copy2(template_path, output_path)
        print("✓ Template copied", flush=True)
        
        # Open the copied template
        print("Opening copied template...", flush=True)
        doc = Document(output_path)
        
        # Clear body content (keep header/footer)
        print("Clearing body content...", flush=True)
        
        for para in doc.paragraphs[:]:
            p_element = para._element
            p_element.getparent().remove(p_element)
        
        for table in doc.tables[:]:
            t_element = table._element
            t_element.getparent().remove(t_element)
        
        print("✓ Body content cleared", flush=True)
        
        # Add new content
        print("\nAdding new content...", flush=True)
        
        # Generate metadata
        today = datetime.now()
        quote_number = f"QT-{today.strftime('%Y-%m%d-%H%M')}"
        quote_date = today.strftime("%B %d, %Y")
        valid_until = (today + timedelta(days=30)).strftime("%B %d, %Y")
        
        # Create helper
        template_helper = Offer3Template()
        template_helper.doc = doc
        
        # 1. Document info
        print("  → Adding document info...", flush=True)
        template_helper.add_document_info_table(quote_number, quote_date, valid_until, "[Customer Name]")
        
        # 2. Pricing table
        print("  → Adding pricing table...", flush=True)
        template_helper.add_pricing_table(items)
        
        # 3. Technical content with STRUCTURE
        print("  → Adding structured technical content...", flush=True)
        add_structured_content_to_doc(doc, items)
        
        # 4. Commercial terms
        print("  → Adding commercial terms...", flush=True)
        template_helper.add_commercial_terms(company_data)
        
        # Save
        print(f"\nSaving document: {output_path}", flush=True)
        doc.save(output_path)
        
        file_size = os.path.getsize(output_path)
        print(f"✓ Document saved: {file_size:,} bytes", flush=True)
        
        # Summary
        print("\n" + "=" * 60, flush=True)
        print("OFFER 3 GENERATION SUMMARY", flush=True)
        print("=" * 60, flush=True)
        print(f"Total Items: {len(items)}", flush=True)
        print(f"Items with structured content: {items_with_blocks}", flush=True)
        print(f"Total content blocks: {total_blocks}", flush=True)
        print(f"Quote Number: {quote_number}", flush=True)
        print(f"Valid Until: {valid_until}", flush=True)
        print("=" * 60, flush=True)
        
        print("\n✓ OFFER 3 GENERATION COMPLETED SUCCESSFULLY", flush=True)
        
        return True
        
    except Exception as e:
        print(f"\n✗ FATAL ERROR: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Offer 3 Generation Script Started", flush=True)
    
    company_data_path = os.path.join(OUTPUT_FOLDER, "company_data.json")
    items_data_path = os.path.join(OUTPUT_FOLDER, "items_offer1.json")
    output_path = os.path.join(OUTPUT_FOLDER, "final_offer3.docx")
    
    success = generate_offer3(company_data_path, items_data_path, output_path)
    
    if not success:
        print("✗ Generation failed", flush=True)
        sys.exit(1)
    
    print("✓ COMPLETED SUCCESSFULLY", flush=True)
    sys.exit(0)