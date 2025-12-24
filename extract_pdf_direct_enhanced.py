"""
TWO-PHASE EXTRACTION - Ensures complete technical content capture

PHASE 1: Extract pricing table structure (item names, prices, quantities)
PHASE 2: Extract ALL technical content as complete sections (no truncation)

This approach prevents AI from getting "lazy" and truncating descriptions.
"""

import os
import sys
import json
import openai
import fitz
import base64
import time

openai.api_key = os.environ.get('OPENAI_API_KEY')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Configuration
MAX_PAGES = 15
IMAGE_SCALE = 1.5
MAX_TOKENS_PRICING = 4000
MAX_TOKENS_TECHNICAL = 8000

def extract_items_from_pdf(pdf_path, output_path):
    try:
        print("=" * 80, flush=True)
        print("TWO-PHASE EXTRACTION - Pricing + Complete Technical Content", flush=True)
        print("=" * 80, flush=True)
        
        start_time = time.time()
        
        if not openai.api_key:
            print("ERROR: OPENAI_API_KEY not set", flush=True)
            return False
        
        print("Reading PDF:", pdf_path, flush=True)
        
        if not os.path.exists(pdf_path):
            print("ERROR: PDF file not found", flush=True)
            return False
        
        # Convert PDF pages to images
        print("Converting PDF to images...", flush=True)
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        print(f"PDF has {total_pages} pages", flush=True)
        
        max_pages = min(MAX_PAGES, total_pages)
        print(f"Processing first {max_pages} pages", flush=True)
        
        image_data_list = []
        for page_num in range(max_pages):
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(IMAGE_SCALE, IMAGE_SCALE))
            img_bytes = pix.tobytes("png")
            img_base64 = base64.b64encode(img_bytes).decode('utf-8')
            image_data_list.append(f"data:image/png;base64,{img_base64}")
            print(f"  Page {page_num + 1}: converted", flush=True)
        
        doc.close()
        print(f"✓ All pages converted ({time.time() - start_time:.1f}s)", flush=True)
        
        # =================================================================
        # PHASE 1: EXTRACT PRICING TABLE STRUCTURE
        # =================================================================
        print("\n" + "=" * 80, flush=True)
        print("PHASE 1: EXTRACTING PRICING TABLE STRUCTURE", flush=True)
        print("=" * 80, flush=True)
        
        pricing_content = [
            {"type": "text", "text": """Extract the PRICING TABLE STRUCTURE ONLY.

Your job: Find all priced items and extract basic table data.

WHAT TO EXTRACT:
For each item in pricing tables, extract:
- category: Section name (e.g., "Main Equipment", "Options", "Packing")
- item_name: Short name from table (e.g., "MODULAR CM 576-9-SM-4B 2-0-0-0-0")
- quantity: Number (default "1" if not shown)
- unit_price: Keep EXACT format:
  * Numeric: "€150.320,00" (preserve dots/commas)
  * Included: "Included"
  * On request: "On request"
- total_price: Same format as unit_price

WHAT NOT TO EXTRACT:
- Do NOT extract technical descriptions
- Do NOT extract specifications
- Do NOT extract detailed features
- ONLY extract what's IN THE TABLE CELLS

SECTIONS TO LOOK FOR:
- Main Equipment
- Options / Accessories
- Format Changes
- Spare Parts
- Packing / Transportation
- Any other priced sections

Return ONLY valid JSON array:
[{
  "category": "Main Equipment",
  "item_name": "MODULAR CM 576-9-SM-4B 2-0-0-0-0",
  "quantity": "1",
  "unit_price": "€150.320,00",
  "total_price": "€150.320,00"
}]

No markdown, no explanations, just JSON array.
"""}
        ]
        
        for img_data in image_data_list:
            pricing_content.append({"type": "image_url", "image_url": {"url": img_data}})
        
        print("Calling GPT-4o for pricing extraction...", flush=True)
        phase1_start = time.time()
        
        response_pricing = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": pricing_content}],
            max_tokens=MAX_TOKENS_PRICING,
            temperature=0
        )
        
        print(f"✓ Phase 1 completed ({time.time() - phase1_start:.1f}s)", flush=True)
        
        pricing_json = response_pricing.choices[0].message.content.strip()
        
        # Clean JSON
        if pricing_json.startswith("```json"):
            pricing_json = pricing_json.replace("```json", "").replace("```", "").strip()
        elif pricing_json.startswith("```"):
            pricing_json = pricing_json.replace("```", "").strip()
        
        items = json.loads(pricing_json)
        print(f"✓ Extracted {len(items)} pricing items", flush=True)
        
        # Show categories
        categories = {}
        for item in items:
            cat = item.get("category", "Unknown")
            categories[cat] = categories.get(cat, 0) + 1
        
        print("Categories found:", flush=True)
        for cat, count in categories.items():
            print(f"  - {cat}: {count} items", flush=True)
        
        # =================================================================
        # PHASE 2: EXTRACT COMPLETE TECHNICAL CONTENT
        # =================================================================
        print("\n" + "=" * 80, flush=True)
        print("PHASE 2: EXTRACTING COMPLETE TECHNICAL CONTENT", flush=True)
        print("=" * 80, flush=True)
        
        technical_content = [
            {"type": "text", "text": """Extract ALL TECHNICAL CONTENT from this document.

Your job: Find and extract COMPLETE technical descriptions, specifications, and details.

WHAT TO EXTRACT:
Find content that appears AFTER the pricing table but BEFORE commercial terms.
This includes:
- Product descriptions and overviews
- Technical specifications sections
- Feature descriptions
- Operating principles
- Materials and construction details
- Standards and certifications
- Any technical paragraphs related to the products

HOW TO EXTRACT:
1. Look for numbered sections (e.g., "1. MODULAR CM...", "2. Side-mounted air...")
2. For EACH section, extract:
   - section_number: The number (e.g., "1", "2", "3")
   - heading: The section title
   - full_content: COMPLETE text - do NOT summarize, do NOT truncate
   
3. If there's a "Key Specifications" subsection, extract it separately

CRITICAL RULES:
- Extract VERBATIM - copy text word-for-word
- Do NOT shorten or summarize
- Do NOT skip any paragraphs
- Do NOT stop after first few sections
- Continue until you see "Commercial Terms" or "Payment Terms"
- If content is 500+ words, that's GOOD - keep it all

EXAMPLE OUTPUT:
[{
  "section_number": "1",
  "heading": "MODULAR CM 576-9-SM-4B 2-0-0-0-0",
  "full_content": "The MODULAR CM 576-9-SM-4B 2-0-0-0-0 is an automatic rotary labelling machine designed with a new construction concept for easy and quick reconfiguration. It features a modular and ergonomic design with a rounded-steel frame towards the guards. The machine-carrying structures are made of AISI 304 stainless steel, while the lower one is steel. It includes a bottle transfer system with anti-wear and tear plastic material components. The bottle carousel has a bottle plate covered with AISI 304 and can be equipped with centering or spotting plates. The machine offers speed adjustment and automatic control, with a pneumatic system ready to connect to existing lines. It complies with EC standards and includes manual greasing or grouped grease nipples. The machine is controlled by a PLC with a touch-screen control panel, automatic speed variation, and safety guards according to EEC rules.",
  "specifications": "Label length: 12-140mm, Label height: 20-180mm, Container diameter: 50-120mm, Container height: 110-350mm, Alternated bottle height: 210-440mm, Machine diameter: 576mm, Machine height: 2050-2310mm, Output: 3000 bph, Machine pitch: 200-400mm, Starwheels diameter: 384mm, No. starwheel divisions: 6, Voltage: 400V (+/-10%) 3F + N + PE, Controls: 24V cc, Material: AISI 304 stainless steel, Standards: EC certified"
},
{
  "section_number": "2",
  "heading": "Side-mounted air conditioning for electrical cabinet",
  "full_content": "Side-mounted air conditioning unit for electrical cabinet, rated at IP55. Designed to maintain optimal temperature conditions within the electrical cabinet, ensuring reliable operation of the machine's electronic components.",
  "specifications": ""
}]

Return ONLY valid JSON array. No markdown, no explanations.
"""}
        ]
        
        for img_data in image_data_list:
            technical_content.append({"type": "image_url", "image_url": {"url": img_data}})
        
        print("Calling GPT-4o for technical content extraction...", flush=True)
        phase2_start = time.time()
        
        response_technical = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": technical_content}],
            max_tokens=MAX_TOKENS_TECHNICAL,
            temperature=0
        )
        
        print(f"✓ Phase 2 completed ({time.time() - phase2_start:.1f}s)", flush=True)
        
        technical_json = response_technical.choices[0].message.content.strip()
        
        # Clean JSON
        if technical_json.startswith("```json"):
            technical_json = technical_json.replace("```json", "").replace("```", "").strip()
        elif technical_json.startswith("```"):
            technical_json = technical_json.replace("```", "").strip()
        
        technical_sections = json.loads(technical_json)
        print(f"✓ Extracted {len(technical_sections)} technical sections", flush=True)
        
        # Quality check
        total_chars = sum(len(section.get('full_content', '')) for section in technical_sections)
        avg_chars = total_chars // len(technical_sections) if technical_sections else 0
        
        print(f"Quality metrics:", flush=True)
        print(f"  Total technical content: {total_chars:,} characters", flush=True)
        print(f"  Average per section: {avg_chars} characters", flush=True)
        
        # Show sample
        if technical_sections:
            print("\nSample technical section:", flush=True)
            sample = technical_sections[0]
            print(f"  Section: {sample.get('heading', 'N/A')}", flush=True)
            content = sample.get('full_content', '')
            print(f"  Content length: {len(content)} chars", flush=True)
            print(f"  First 150 chars: {content[:150]}...", flush=True)
        
        # =================================================================
        # MERGE PHASES: Match technical content to pricing items
        # =================================================================
        print("\n" + "=" * 80, flush=True)
        print("MERGING: Matching technical content to pricing items", flush=True)
        print("=" * 80, flush=True)
        
        # Create mapping by section number
        tech_by_number = {}
        for section in technical_sections:
            num = section.get('section_number', '')
            if num:
                tech_by_number[num] = section
        
        # Match to items
        matched = 0
        for idx, item in enumerate(items):
            item_number = str(idx + 1)
            
            if item_number in tech_by_number:
                tech = tech_by_number[item_number]
                item['description'] = tech.get('full_content', '')
                item['specifications'] = tech.get('specifications', '')
                item['details'] = ''
                matched += 1
            else:
                # No technical content for this item
                item['description'] = ''
                item['specifications'] = ''
                item['details'] = ''
        
        print(f"✓ Matched {matched}/{len(items)} items with technical content", flush=True)
        
        if matched < len(items):
            print(f"⚠ {len(items) - matched} items have no technical content", flush=True)
        
        # =================================================================
        # SAVE OUTPUT
        # =================================================================
        output_data = {
            "items": items,
            "technical_sections": technical_sections,
            "categories": list(categories.keys()),
            "extraction_version": "SV16_TwoPhase",
            "extraction_stats": {
                "total_items": len(items),
                "items_with_descriptions": matched,
                "total_technical_chars": total_chars,
                "avg_chars_per_section": avg_chars
            }
        }
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False)
        
        total_time = time.time() - start_time
        
        print("\n" + "=" * 80, flush=True)
        print("EXTRACTION COMPLETED SUCCESSFULLY", flush=True)
        print("=" * 80, flush=True)
        print(f"Total time: {total_time:.1f}s", flush=True)
        print(f"Output: {output_path}", flush=True)
        print(f"Items: {len(items)}", flush=True)
        print(f"Technical sections: {len(technical_sections)}", flush=True)
        print(f"Match rate: {matched}/{len(items)} ({matched*100//len(items) if items else 0}%)", flush=True)
        print("=" * 80, flush=True)
        
        return True
        
    except Exception as e:
        print(f"\n✗ FATAL ERROR: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Two-Phase Extraction Script Started", flush=True)
    pdf_path = os.path.join(UPLOAD_FOLDER, "offer1.pdf")
    output_path = os.path.join(OUTPUT_FOLDER, "items_offer1.json")
    
    if not os.path.exists(pdf_path):
        print("ERROR: PDF not found at " + pdf_path, flush=True)
        sys.exit(1)
    
    success = extract_items_from_pdf(pdf_path, output_path)
    
    if not success:
        print("Extraction failed", flush=True)
        sys.exit(1)
    
    print("COMPLETED SUCCESSFULLY", flush=True)
    sys.exit(0)