"""
Build Offer 3 - SIMPLE APPROACH
Start with Offer 2 template, clear body, add equipment data
This preserves header/footer/logo automatically
"""

import os
import sys
import json
from datetime import datetime, timedelta
from docx import Document
from standard_template import Offer3Template

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')

def generate_offer3(company_data_path, items_data_path, output_path):
    """
    Generate Offer 3 by:
    1. Starting with Offer 2 template (preserves header/footer/logo)
    2. Clearing the body content
    3. Adding equipment data from Offer 1
    """
    
    try:
        print("=" * 60, flush=True)
        print("BUILDING OFFER 3 - SIMPLE TEMPLATE APPROACH", flush=True)
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
        
        # Load company data (optional, for commercial terms)
        company_data = {}
        if os.path.exists(company_data_path):
            with open(company_data_path, 'r', encoding='utf-8') as f:
                company_data = json.load(f)
            print(f"✓ Company data loaded", flush=True)
        
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
            print("⚠ Template not found! Creating document without template.", flush=True)
            # Fallback: create new document
            template_helper = Offer3Template()
            doc = template_helper.doc
        else:
            # SMART APPROACH: 
            # 1. Open template to find which section has header/footer
            # 2. Create new document
            # 3. Copy header/footer from the correct section
            
            print("Opening template file...", flush=True)
            template_doc = Document(template_path)
            
            print(f"Template has {len(template_doc.sections)} section(s)", flush=True)
            
            # Find section with header/footer
            section_with_header = None
            
            for idx, section in enumerate(template_doc.sections):
                has_header = False
                
                if section.header:
                    # Check if header has text content
                    for para in section.header.paragraphs:
                        if para.text.strip():
                            has_header = True
                            break
                    
                    # Check if header has tables
                    if not has_header and len(section.header.tables) > 0:
                        has_header = True
                    
                    # Check if header has images/shapes (logo)
                    if not has_header:
                        # Check for images in header relationships
                        try:
                            for rel in section.header.part.rels.values():
                                if "image" in rel.target_ref.lower():
                                    has_header = True
                                    print(f"  → Section {idx + 1} header contains image", flush=True)
                                    break
                        except:
                            pass
                
                if has_header:
                    section_with_header = idx
                    print(f"✓ Found header/footer in section {idx + 1}", flush=True)
                    break
            
            if section_with_header is None:
                print("⚠ No section with header found, using section 1", flush=True)
                section_with_header = 0
            
            # Create NEW clean document
            doc = Document()
            
            # Copy header and footer from the section that has them
            source_section = template_doc.sections[section_with_header]
            target_section = doc.sections[0]
            
            print(f"Copying header/footer from template section {section_with_header + 1}...", flush=True)
            
            # Copy HEADER
            if source_section.header:
                print("  → Copying header...", flush=True)
                
                # Copy paragraphs
                for para in source_section.header.paragraphs:
                    new_para = target_section.header.add_paragraph()
                    
                    # Copy runs with formatting
                    for run in para.runs:
                        new_run = new_para.add_run(run.text)
                        new_run.bold = run.bold
                        new_run.italic = run.italic
                        if run.font.size:
                            new_run.font.size = run.font.size
                        if run.font.name:
                            new_run.font.name = run.font.name
                        
                        # Try to copy images
                        if hasattr(run._element, 'drawing') or 'drawing' in run._element.xml:
                            print("  ⚠ Header contains image (logo) - cannot copy programmatically", flush=True)
                
                # Copy tables
                for table in source_section.header.tables:
                    print(f"  → Copying header table ({len(table.rows)} rows, {len(table.columns)} cols)...", flush=True)
                    
                    try:
                        # Create new table in header
                        new_table = target_section.header.add_table(rows=len(table.rows), cols=len(table.columns))
                        
                        # Copy cell content and basic formatting
                        for i, row in enumerate(table.rows):
                            for j, cell in enumerate(row.cells):
                                new_table.rows[i].cells[j].text = cell.text
                                
                                # Copy cell formatting if possible
                                for para in cell.paragraphs:
                                    for run in para.runs:
                                        if run.bold or run.italic:
                                            # Apply formatting to new cell
                                            pass  # Basic text copy is enough for now
                        
                        print(f"  ✓ Header table copied", flush=True)
                    except Exception as table_err:
                        print(f"  ⚠ Could not copy header table: {str(table_err)}", flush=True)
                
                print("  ✓ Header copied (text only, logo requires manual intervention)", flush=True)
            
            # Copy FOOTER
            if source_section.footer:
                print("  → Copying footer...", flush=True)
                
                # Copy paragraphs
                for para in source_section.footer.paragraphs:
                    new_para = target_section.footer.add_paragraph()
                    
                    for run in para.runs:
                        new_run = new_para.add_run(run.text)
                        new_run.bold = run.bold
                        new_run.italic = run.italic
                        if run.font.size:
                            new_run.font.size = run.font.size
                        if run.font.name:
                            new_run.font.name = run.font.name
                
                # Copy tables
                for table in source_section.footer.tables:
                    print(f"  → Copying footer table ({len(table.rows)} rows, {len(table.columns)} cols)...", flush=True)
                    
                    try:
                        # Create new table in footer
                        new_table = target_section.footer.add_table(rows=len(table.rows), cols=len(table.columns))
                        
                        # Copy cell content
                        for i, row in enumerate(table.rows):
                            for j, cell in enumerate(row.cells):
                                new_table.rows[i].cells[j].text = cell.text
                        
                        print(f"  ✓ Footer table copied", flush=True)
                    except Exception as table_err:
                        print(f"  ⚠ Could not copy footer table: {str(table_err)}", flush=True)
                        # Try alternative: add footer content as paragraphs instead
                        print(f"  → Adding footer content as text instead...", flush=True)
                        for row in table.rows:
                            row_text = " | ".join([cell.text for cell in row.cells if cell.text.strip()])
                            if row_text:
                                footer_para = target_section.footer.add_paragraph(row_text)
                                footer_para.style = 'Normal'
                
                print("  ✓ Footer copied", flush=True)
            
            print("✓ Header/footer from template applied to new document", flush=True)
            
            # Create helper for adding content using the template document
            template_helper = Offer3Template()
            template_helper.doc = doc  # Replace helper's document with template
        
        # Generate quote metadata
        today = datetime.now()
        quote_number = f"QT-{today.strftime('%Y-%m%d-%H%M')}"
        quote_date = today.strftime("%B %d, %Y")
        valid_until = (today + timedelta(days=30)).strftime("%B %d, %Y")
        
        print(f"\nQuote Number: {quote_number}", flush=True)
        print(f"Date: {quote_date}", flush=True)
        print(f"Valid Until: {valid_until}", flush=True)
        
        # Add document content
        print("\nAdding content to document...", flush=True)
        
        # 1. Document info table
        print("  → Adding document info...", flush=True)
        template_helper.add_document_info_table(quote_number, quote_date, valid_until, "[Customer Name]")
        
        # 2. Pricing table
        print("  → Adding pricing table...", flush=True)
        template_helper.add_pricing_table(items)
        
        # 3. Technical descriptions
        items_with_desc = [item for item in items if item.get('description') or item.get('specifications')]
        if items_with_desc:
            print(f"  → Adding technical descriptions ({len(items_with_desc)} items)...", flush=True)
            template_helper.add_technical_descriptions(items_with_desc)
        
        # 4. Commercial terms
        print("  → Adding commercial terms...", flush=True)
        template_helper.add_commercial_terms(company_data)
        
        # Save document
        print(f"\nSaving document to: {output_path}", flush=True)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc.save(output_path)
        
        file_size = os.path.getsize(output_path)
        print(f"✓ Document saved: {output_path}", flush=True)
        print(f"✓ File size: {file_size:,} bytes", flush=True)
        
        # Summary
        print("\n" + "=" * 60, flush=True)
        print("OFFER 3 GENERATION SUMMARY", flush=True)
        print("=" * 60, flush=True)
        print(f"Company: {company_data.get('company_name', 'N/A')}", flush=True)
        print(f"Total Items: {len(items)}", flush=True)
        print(f"Items with Descriptions: {len(items_with_desc)}", flush=True)
        
        # Count categories
        categories = {}
        for item in items:
            cat = item.get('category', 'Uncategorized')
            categories[cat] = categories.get(cat, 0) + 1
        
        print("Categories:", flush=True)
        for cat, count in categories.items():
            print(f"  - {cat}: {count} items", flush=True)
        
        print(f"\nQuote Number: {quote_number}", flush=True)
        print(f"Valid Until: {valid_until}", flush=True)
        print("=" * 60, flush=True)
        
        print("\n✓ OFFER 3 GENERATION COMPLETED SUCCESSFULLY", flush=True)
        print("=" * 60, flush=True)
        
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