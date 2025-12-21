"""
Build Offer 3 - ACTUALLY SIMPLE APPROACH
Just copy the template file and replace its body content
This preserves EVERYTHING: header, footer, logo, formatting
"""

import os
import sys
import json
import shutil
from datetime import datetime, timedelta
from docx import Document
from standard_template import Offer3Template

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')

def generate_offer3(company_data_path, items_data_path, output_path):
    """
    SIMPLE APPROACH:
    1. Copy template file to output
    2. Open it
    3. Clear body
    4. Add new content
    5. Save
    
    Header/footer/logo preserved automatically because we're editing the template itself.
    """
    
    try:
        print("=" * 60, flush=True)
        print("BUILDING OFFER 3 - COPY AND EDIT TEMPLATE", flush=True)
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
            print("✗ Template not found! Cannot create offer without template.", flush=True)
            return False
        
        # STEP 1: Copy template to output location
        print(f"Copying template to: {output_path}", flush=True)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        shutil.copy2(template_path, output_path)
        print("✓ Template copied", flush=True)
        
        # STEP 2: Open the copied template
        print("Opening copied template...", flush=True)
        doc = Document(output_path)
        
        print(f"Template has {len(doc.sections)} section(s)", flush=True)
        print(f"Template has {len(doc.paragraphs)} paragraph(s)", flush=True)
        print(f"Template has {len(doc.tables)} table(s)", flush=True)
        
        # STEP 3: Delete all body content (keep header/footer)
        print("Clearing body content...", flush=True)
        
        # Remove all paragraphs
        for para in doc.paragraphs[:]:
            p_element = para._element
            p_element.getparent().remove(p_element)
        
        # Remove all tables
        for table in doc.tables[:]:
            t_element = table._element
            t_element.getparent().remove(t_element)
        
        print("✓ Body content cleared", flush=True)
        
        # STEP 4: Add new content
        print("\nAdding new content...", flush=True)
        
        # Generate metadata
        today = datetime.now()
        quote_number = f"QT-{today.strftime('%Y-%m%d-%H%M')}"
        quote_date = today.strftime("%B %d, %Y")
        valid_until = (today + timedelta(days=30)).strftime("%B %d, %Y")
        
        # Create helper (but use our document, not a new one)
        template_helper = Offer3Template()
        template_helper.doc = doc
        
        # Add document info
        print("  → Adding document info...", flush=True)
        template_helper.add_document_info_table(quote_number, quote_date, valid_until, "[Customer Name]")
        
        # Add pricing table
        print("  → Adding pricing table...", flush=True)
        template_helper.add_pricing_table(items)
        
        # Add technical descriptions
        items_with_desc = [item for item in items if item.get('description') or item.get('specifications')]
        if items_with_desc:
            print(f"  → Adding technical descriptions ({len(items_with_desc)} items)...", flush=True)
            template_helper.add_technical_descriptions(items_with_desc)
        
        # Add commercial terms
        print("  → Adding commercial terms...", flush=True)
        template_helper.add_commercial_terms(company_data)
        
        # STEP 5: Save (overwrites the copied template with new content)
        print(f"\nSaving document: {output_path}", flush=True)
        doc.save(output_path)
        
        file_size = os.path.getsize(output_path)
        print(f"✓ Document saved: {file_size:,} bytes", flush=True)
        
        # Summary
        print("\n" + "=" * 60, flush=True)
        print("OFFER 3 GENERATION SUMMARY", flush=True)
        print("=" * 60, flush=True)
        print(f"Total Items: {len(items)}", flush=True)
        print(f"Items with Descriptions: {len(items_with_desc)}", flush=True)
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