"""
THREE-PHASE EXTRACTION with SEMANTIC MATCHING

PHASE 1: Understand the offer (what's being sold)
PHASE 2: Extract pricing table structure
PHASE 3: Extract ALL technical content with context-aware matching
"""

import os
import sys
import json
import openai
import fitz
import base64
import time
from difflib import SequenceMatcher

openai.api_key = os.environ.get('OPENAI_API_KEY')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Configuration
MAX_PAGES = 15
IMAGE_SCALE = 1.5
MAX_TOKENS_CONTEXT = 2000
MAX_TOKENS_PRICING = 4000
MAX_TOKENS_TECHNICAL = 8000

def similarity(a, b):
    """Calculate similarity between two strings (0-1)"""
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

def extract_items_from_pdf(pdf_path, output_path):
    try:
        print("=" * 80, flush=True)
        print("THREE-PHASE SEMANTIC EXTRACTION", flush=True)
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
        # PHASE 1: UNDERSTAND THE OFFER CONTEXT
        # =================================================================
        print("\n" + "=" * 80, flush=True)
        print("PHASE 1: UNDERSTANDING OFFER CONTEXT", flush=True)
        print("=" * 80, flush=True)
        
        context_content = [
            {"type": "text", "text": """Analyze this commercial offer and provide context.

Answer these questions:
1. What is the MAIN equipment/product being offered?
2. What is the manufacturer/supplier name?
3. What industry is this for? (e.g., beverage, food processing, packaging)

Return ONLY JSON:
{
  "main_product": "Brief description of main equipment",
  "supplier": "Company name",
  "industry": "Industry sector",
  "offer_type": "quotation" or "catalog" or "technical sheet"
}
"""}
        ]
        
        # Use only first 3 pages for context
        for img_data in image_data_list[:3]:
            context_content.append({"type": "image_url", "image_url": {"url": img_data}})
        
        print("Calling GPT-4o for context analysis...", flush=True)
        phase1_start = time.time()
        
        response_context = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": context_content}],
            max_tokens=MAX_TOKENS_CONTEXT,
            temperature=0
        )
        
        print(f"✓ Phase 1 completed ({time.time() - phase1_start:.1f}s)", flush=True)
        
        context_json = response_context.choices[0].message.content.strip()
        if context_json.startswith("```json"):
            context_json = context_json.replace("```json", "").replace("```", "").strip()
        elif context_json.startswith("```"):
            context_json = context_json.replace("```", "").strip()
        
        offer_context = json.loads(context_json)
        print(f"Offer context:", flush=True)
        print(f"  Main product: {offer_context.get('main_product', 'Unknown')}", flush=True)
        print(f"  Supplier: {offer_context.get('supplier', 'Unknown')}", flush=True)
        print(f"  Industry: {offer_context.get('industry', 'Unknown')}", flush=True)
        
        # =================================================================
        # PHASE 2: EXTRACT PRICING TABLE
        # =================================================================
        print("\n" + "=" * 80, flush=True)
        print("PHASE 2: EXTRACTING PRICING TABLE", flush=True)
        print("=" * 80, flush=True)
        
        pricing_content = [
            {"type": "text", "text": """Extract the PRICING TABLE with item identification keys.

For each priced item, extract:
- item_number: Position in table (1, 2, 3, etc.)
- category: Section (e.g., "Main Equipment", "Options", "Packing")
- item_name: FULL name from table (important for matching)
- quantity: Number
- unit_price: Exact format (€X or "Included" or "On request")
- total_price: Same format

CRITICAL: Extract the COMPLETE item_name exactly as written in the table.
This is the key for matching technical descriptions later.

Return ONLY JSON array:
[{
  "item_number": 1,
  "category": "Main Equipment",
  "item_name": "MODULAR CM 576-9-SM-4B 2-0-0-0-0",
  "quantity": "1",
  "unit_price": "€150.320,00",
  "total_price": "€150.320,00"
}]
"""}
        ]
        
        for img_data in image_data_list:
            pricing_content.append({"type": "image_url", "image_url": {"url": img_data}})
        
        print("Calling GPT-4o for pricing extraction...", flush=True)
        phase2_start = time.time()
        
        response_pricing = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": pricing_content}],
            max_tokens=MAX_TOKENS_PRICING,
            temperature=0
        )
        
        print(f"✓ Phase 2 completed ({time.time() - phase2_start:.1f}s)", flush=True)
        
        pricing_json = response_pricing.choices[0].message.content.strip()
        if pricing_json.startswith("```json"):
            pricing_json = pricing_json.replace("```json", "").replace("```", "").strip()
        elif pricing_json.startswith("```"):
            pricing_json = pricing_json.replace("```", "").strip()
        
        items = json.loads(pricing_json)
        print(f"✓ Extracted {len(items)} pricing items", flush=True)
        
        # =================================================================
        # PHASE 3: EXTRACT TECHNICAL CONTENT WITH SEMANTIC TAGS
        # =================================================================
        print("\n" + "=" * 80, flush=True)
        print("PHASE 3: EXTRACTING TECHNICAL CONTENT", flush=True)
        print("=" * 80, flush=True)
        
        # Build item reference list for GPT
        item_reference = "\n".join([
            f"{item['item_number']}. {item['item_name']}"
            for item in items
        ])
        
        technical_content = [
            {"type": "text", "text": f"""Extract ALL TECHNICAL CONTENT and match it to the correct items.

CONTEXT:
This is a quotation for: {offer_context.get('main_product', 'industrial equipment')}

PRICING TABLE ITEMS (for reference):
{item_reference}

YOUR TASK:
Extract all technical descriptions that appear AFTER the pricing table.

For EACH technical section you find:
1. Read the section heading carefully
2. Determine WHICH item from the pricing table it describes
3. Extract the COMPLETE content (do NOT truncate)

Return JSON array where each object has:
{{
  "matched_item_number": [Which item number(s) this describes - can be multiple],
  "heading": "The section title",
  "full_content": "COMPLETE verbatim text - do NOT summarize",
  "specifications": "Any structured specs (separate paragraph)",
  "matching_confidence": "high" or "medium" or "low"
}}

MATCHING RULES:
- If heading contains exact item name → matched_item_number = that item
- If content describes main equipment → matched_item_number = 1
- If heading mentions "option 2" or similar → match to item 2
- If unsure → set matching_confidence = "low"

EXTRACTION RULES:
- Extract VERBATIM - copy complete text word-for-word
- Do NOT shorten or summarize  
- Do NOT skip sections
- Include ALL paragraphs related to each item
- Stop when you see "Commercial Terms" or "Payment Terms"

Example output:
[{{
  "matched_item_number": [1],
  "heading": "1. MODULAR CM 576-9-SM-4B 2-0-0-0-0",
  "full_content": "The MODULAR CM 576-9-SM-4B 2-0-0-0-0 is an automatic rotary labelling machine...[COMPLETE TEXT]",
  "specifications": "Label length: 12-140mm, Label height: 20-180mm...",
  "matching_confidence": "high"
}},
{{
  "matched_item_number": [3],
  "heading": "3. Conveyor motor drive with Rossi motoreducer 0,55kw",
  "full_content": "Automatic rotary labelling machine made with...[COMPLETE TEXT]",
  "specifications": "",
  "matching_confidence": "high"
}}]

Return ONLY valid JSON array.
"""}
        ]
        
        for img_data in image_data_list:
            technical_content.append({"type": "image_url", "image_url": {"url": img_data}})
        
        print("Calling GPT-4o for technical extraction...", flush=True)
        phase3_start = time.time()
        
        response_technical = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": technical_content}],
            max_tokens=MAX_TOKENS_TECHNICAL,
            temperature=0
        )
        
        print(f"✓ Phase 3 completed ({time.time() - phase3_start:.1f}s)", flush=True)
        
        technical_json = response_technical.choices[0].message.content.strip()
        if technical_json.startswith("```json"):
            technical_json = technical_json.replace("```json", "").replace("```", "").strip()
        elif technical_json.startswith("```"):
            technical_json = technical_json.replace("```", "").strip()
        
        technical_sections = json.loads(technical_json)
        print(f"✓ Extracted {len(technical_sections)} technical sections", flush=True)
        
        # =================================================================
        # SMART MATCHING: Assign technical content to items
        # =================================================================
        print("\n" + "=" * 80, flush=True)
        print("SEMANTIC MATCHING: Assigning technical content to items", flush=True)
        print("=" * 80, flush=True)
        
        # Initialize empty descriptions
        for item in items:
            item['description'] = ''
            item['specifications'] = ''
            item['details'] = ''
            item['matched_sections'] = []
        
        # Assign based on GPT's matching
        for section in technical_sections:
            matched_nums = section.get('matched_item_number', [])
            if not isinstance(matched_nums, list):
                matched_nums = [matched_nums]
            
            confidence = section.get('matching_confidence', 'low')
            content = section.get('full_content', '')
            specs = section.get('specifications', '')
            
            for num in matched_nums:
                # Find item by number
                for item in items:
                    if item['item_number'] == num:
                        # Append content (multiple sections can describe same item)
                        if item['description']:
                            item['description'] += "\n\n" + content
                        else:
                            item['description'] = content
                        
                        if specs:
                            if item['specifications']:
                                item['specifications'] += " " + specs
                            else:
                                item['specifications'] = specs
                        
                        item['matched_sections'].append({
                            'heading': section.get('heading', ''),
                            'confidence': confidence
                        })
                        
                        print(f"  ✓ Matched section '{section.get('heading', 'Unknown')[:50]}...' to item {num} ({confidence} confidence)", flush=True)
        
        # Report matching statistics
        matched_count = sum(1 for item in items if item['description'])
        print(f"\n✓ Successfully matched {matched_count}/{len(items)} items", flush=True)
        
        # Show unmatched items
        unmatched = [item for item in items if not item['description']]
        if unmatched:
            print(f"\n⚠ {len(unmatched)} items without technical descriptions:", flush=True)
            for item in unmatched:
                print(f"  - Item {item['item_number']}: {item['item_name']}", flush=True)
        
        # =================================================================
        # SAVE OUTPUT
        # =================================================================
        output_data = {
            "offer_context": offer_context,
            "items": items,
            "technical_sections_raw": technical_sections,
            "categories": list(set(item['category'] for item in items)),
            "extraction_version": "SV17_SemanticMatching",
            "extraction_stats": {
                "total_items": len(items),
                "items_with_descriptions": matched_count,
                "unmatched_items": len(unmatched),
                "total_technical_sections": len(technical_sections)
            }
        }
        
        # Clean up matched_sections before saving (optional)
        for item in items:
            if 'matched_sections' in item:
                del item['matched_sections']
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False)
        
        total_time = time.time() - start_time
        
        print("\n" + "=" * 80, flush=True)
        print("EXTRACTION COMPLETED SUCCESSFULLY", flush=True)
        print("=" * 80, flush=True)
        print(f"Total time: {total_time:.1f}s", flush=True)
        print(f"Output: {output_path}", flush=True)
        print(f"Match rate: {matched_count}/{len(items)} ({matched_count*100//len(items) if items else 0}%)", flush=True)
        print("=" * 80, flush=True)
        
        return True
        
    except Exception as e:
        print(f"\n✗ FATAL ERROR: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Three-Phase Semantic Extraction Script Started", flush=True)
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