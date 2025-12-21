"""
Build Offer 3 - Generate new quotation document from scratch
Combines company branding (Offer 2) with equipment data (Offer 1)
"""

import os
import sys
import json
from datetime import datetime, timedelta
from standard_template import Offer3Template

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')

def generate_offer3(company_data_path, items_data_path, output_path):
    """
    Generate Offer 3 by combining:
    - Company data from Offer 2 (branding, terms, bank details)
    - Equipment data from Offer 1 (items, prices, descriptions)
    """
    
    try:
        print("=" * 60, flush=True)
        print("BUILDING OFFER 3 - NEW GENERATION APPROACH", flush=True)
        print("=" * 60, flush=True)
        
        # Load company data
        print(f"Loading company data: {company_data_path}", flush=True)
        if not os.path.exists(company_data_path):
            print("✗ Company data not found. Please process Offer 2 first.", flush=True)
            return False
        
        with open(company_data_path, 'r', encoding='utf-8') as f:
            company_data = json.load(f)
        
        # Validate company data
        company_name = company_data.get('company_name', '')
        if not company_name or company_name.strip() == '':
            print("⚠ WARNING: Company name is empty!", flush=True)
            print("⚠ Company data extraction may have failed", flush=True)
            print(f"⚠ Company data content: {json.dumps(company_data, indent=2)[:500]}", flush=True)
        else:
            print(f"✓ Company: {company_name}", flush=True)
        
        # Check other critical fields
        if company_data.get('address'):
            print(f"✓ Address found: {company_data['address'][:50]}...", flush=True)
        else:
            print("⚠ No address found", flush=True)
        
        if company_data.get('logo'):
            print(f"✓ Logo found: {company_data['logo'].get('size', 0)} bytes", flush=True)
        else:
            print("⚠ No logo found", flush=True)
        
        if company_data.get('bank_details', {}).get('iban'):
            print(f"✓ Bank details found: IBAN {company_data['bank_details']['iban']}", flush=True)
        else:
            print("⚠ No bank details found", flush=True)
        
        # Load items data
        print(f"Loading items data: {items_data_path}", flush=True)
        if not os.path.exists(items_data_path):
            print("✗ Items data not found. Please process Offer 1 first.", flush=True)
            return False
        
        with open(items_data_path, 'r', encoding='utf-8') as f:
            items_data = json.load(f)
        
        items = items_data.get('items', [])
        print(f"✓ Loaded {len(items)} items", flush=True)
        
        if len(items) == 0:
            print("✗ No items to quote", flush=True)
            return False
        
        # Generate document metadata
        today = datetime.now()
        quote_date = today.strftime("%B %d, %Y")
        valid_until = (today + timedelta(days=30)).strftime("%B %d, %Y")
        quote_number = today.strftime("QT-%Y-%m%d-%H%M")
        
        print(f"Quote Number: {quote_number}", flush=True)
        print(f"Date: {quote_date}", flush=True)
        print(f"Valid Until: {valid_until}", flush=True)
        
        # Create template instance
        print("\nBuilding document...", flush=True)
        template = Offer3Template()
        
        # 1. Copy header and footer directly from Offer 2 template
        print("  → Copying header and footer from template...", flush=True)
        template_path = os.path.join(BASE_DIR, 'offer2_template.docx')
        if os.path.exists(template_path):
            template.copy_header_footer_from_template(template_path)
        else:
            print("  ⚠ Template not found, skipping header/footer copy", flush=True)
        
        # 2. Add document info table
        print("  → Adding document info...", flush=True)
        template.add_document_info_table(
            quote_number=quote_number,
            date=quote_date,
            valid_until=valid_until,
            customer_name="[Customer Name]"  # Placeholder
        )
        
        # 3. Add pricing table
        print("  → Adding pricing table...", flush=True)
        template.add_pricing_table(items)
        
        # 4. Add technical descriptions (only for items with descriptions)
        print("  → Adding technical descriptions...", flush=True)
        items_with_descriptions = [
            item for item in items 
            if item.get('description') or item.get('specifications') or item.get('details')
        ]
        print(f"    ({len(items_with_descriptions)} items have technical content)", flush=True)
        
        if items_with_descriptions:
            template.add_technical_descriptions(items)
        
        # 5. Add commercial terms
        print("  → Adding commercial terms...", flush=True)
        template.add_commercial_terms(company_data)
        
        # 6. Add footer
        print("  → Adding footer...", flush=True)
        template.add_footer_section(company_data)
        
        # Save document
        print(f"\nSaving document to: {output_path}", flush=True)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        template.save(output_path)
        
        # Verification
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print(f"✓ File size: {file_size:,} bytes", flush=True)
            
            # Quality metrics
            print("\n" + "=" * 60, flush=True)
            print("OFFER 3 GENERATION SUMMARY", flush=True)
            print("=" * 60, flush=True)
            print(f"Company: {company_data.get('company_name', 'N/A')}", flush=True)
            print(f"Total Items: {len(items)}", flush=True)
            print(f"Items with Descriptions: {len(items_with_descriptions)}", flush=True)
            
            # Count categories
            categories = set(item.get('category', 'Items') for item in items)
            print(f"Categories: {len(categories)}", flush=True)
            for cat in categories:
                cat_count = sum(1 for item in items if item.get('category') == cat)
                print(f"  - {cat}: {cat_count} items", flush=True)
            
            print(f"\nQuote Number: {quote_number}", flush=True)
            print(f"Valid Until: {valid_until}", flush=True)
            print("=" * 60, flush=True)
            
            return True
        else:
            print("✗ File not created", flush=True)
            return False
        
    except Exception as e:
        print(f"✗ FATAL ERROR: {str(e)}", flush=True)
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
    
    print("\n✓ OFFER 3 GENERATION COMPLETED SUCCESSFULLY", flush=True)
    sys.exit(0)