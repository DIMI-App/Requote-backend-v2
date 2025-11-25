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
MAX_PAGES = 15  # Process max 15 pages
IMAGE_SCALE = 1.5  # Lower resolution for faster processing (was 2.0)
MAX_TOKENS = 8000  # Max tokens for response

def extract_items_from_pdf(pdf_path, output_path):
    try:
        print("=== STARTING ENHANCED EXTRACTION (Prices + Technical Descriptions) ===", flush=True)
        start_time = time.time()
        
        if not openai.api_key:
            print("ERROR: OPENAI_API_KEY not set", flush=True)
            return False
        
        print("OpenAI key found", flush=True)
        print(f"Reading PDF: {pdf_path}", flush=True)
        
        if not os.path.exists(pdf_path):
            print("ERROR: PDF file not found", flush=True)
            return False
        
        print("Opening PDF with PyMuPDF...", flush=True)
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        print(f"PDF has {total_pages} pages", flush=True)
        
        # Process pages with limit
        max_pages = min(MAX_PAGES, total_pages)
        print(f"Processing first {max_pages} pages (scale: {IMAGE_SCALE}x)", flush=True)
        
        image_data_list = []
        for page_num in range(max_pages):
            page_start = time.time()
            print(f"Converting page {page_num + 1}/{max_pages}...", flush=True)
            
            page = doc[page_num]
            # Use configurable scale for balance between quality and speed
            pix = page.get_pixmap(matrix=fitz.Matrix(IMAGE_SCALE, IMAGE_SCALE))
            img_bytes = pix.tobytes("png")
            img_base64 = base64.b64encode(img_bytes).decode('utf-8')
            image_data_list.append(f"data:image/png;base64,{img_base64}")
            
            page_time = time.time() - page_start
            print(f"Page {page_num + 1}: converted in {page_time:.1f}s ({len(img_base64)} bytes)", flush=True)
        
        doc.close()
        conversion_time = time.time() - start_time
        print(f"All pages converted in {conversion_time:.1f}s", flush=True)
        
        print("Building OpenAI request for DUAL extraction (prices + technical content)...", flush=True)
        
        content = [
            {"type": "text", "text": """Extract BOTH pricing items AND technical descriptions from this quotation document.

CRITICAL RULES:
1. Extract from ALL PAGES - scan entire document
2. Extract TWO types of content:

=== PART A: PRICED ITEMS ===
Extract items from these sections:
   - Main equipment/machinery
   - Economic Offer table (main pricing table)
   - Format Changes
   - Accessories sections
   - Further Options
   - Packing options

For each item, capture:
   - category: Section name (e.g., "Main Equipment", "ACCESSORIES", "FORMAT CHANGES")
   - item_name: Item description (full text)
   - quantity: Quantity (default "1" if not shown)
   - unit_price: Use ONE of these THREE states:
     * Numeric price: "€324.400,00" (keep exact format with dots/commas)
     * Included: "Included" (when text says "Included" or price is 0)
     * To be quoted: "On request" (when text says "Can be offered", "To be quoted", etc.)
   - total_price: Same format as unit_price
   - details: Technical specifications, model numbers, features

=== PART B: TECHNICAL DESCRIPTIONS ===
Extract all technical content that is NOT in pricing tables:
   - Product overview/introduction paragraphs
   - Feature descriptions (e.g., "Features of the rinsing turret")
   - Technical specification sections
   - Operation descriptions
   - Equipment capabilities
   - Configuration details

For each technical section:
   - section_title: Title or heading of the section
   - content_type: "paragraph" | "bullet_list" | "spec_table" | "features"
   - content: Full text content (max 500 chars per section)
   - page_location: "before_price_table" | "after_price_table"

IMPORTANT:
- Extract technical content that appears OUTSIDE of pricing tables
- Keep technical descriptions concise (max 500 chars each)
- Continue extracting until you see "GENERAL SALE TERMS" or end of document

Return ONLY JSON:
{
  "items": [
    {
      "category": "Main Equipment",
      "item_name": "CAN ISO 20/2 S",
      "quantity": "1",
      "unit_price": "€324.400,00",
      "total_price": "€324.400,00",
      "details": "Based on 0,33L standard aluminium can"
    }
  ],
  "technical_sections": [
    {
      "section_title": "Features of the rinsing turret",
      "content_type": "features",
      "content": "The rinsing turret is equipped with...",
      "page_location": "before_price_table"
    }
  ]
}
"""}
        ]
        
        for img_data in image_data_list:
            content.append({"type": "image_url", "image_url": {"url": img_data}})
        
        print("Calling OpenAI Vision API (gpt-4o)...", flush=True)
        api_start = time.time()
        
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": content}],
            max_tokens=MAX_TOKENS,
            temperature=0
        )
        
        api_time = time.time() - api_start
        print(f"Received response from OpenAI in {api_time:.1f}s", flush=True)
        
        extracted_json = response.choices[0].message.content.strip()
        
        if extracted_json.startswith("```json"):
            extracted_json = extracted_json.replace("```json", "").replace("```", "").strip()
        elif extracted_json.startswith("```"):
            extracted_json = extracted_json.replace("```", "").strip()
        
        print("Parsing JSON...", flush=True)
        full_data = json.loads(extracted_json)
        
        items = full_data.get("items", [])
        technical_sections = full_data.get("technical_sections", [])
        
        print(f"✓ Extracted {len(items)} items", flush=True)
        print(f"✓ Extracted {len(technical_sections)} technical sections", flush=True)
        
        if len(items) == 0:
            print("WARNING: No items extracted", flush=True)
        
        if len(technical_sections) > 0:
            print(f"Technical sections found:", flush=True)
            for section in technical_sections[:3]:
                print(f"  - {section.get('section_title', 'Untitled')}: {section.get('content_type', 'unknown')}", flush=True)
        
        # Group items by category
        categories = {}
        for item in items:
            cat = item.get("category", "Main Items")
            if cat not in categories:
                categories[cat] = []
            categories[cat].append(item)
        
        print(f"Found {len(categories)} categories:", flush=True)
        for cat, cat_items in categories.items():
            print(f"  - {cat}: {len(cat_items)} items", flush=True)
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # Save full data
        output_data = {
            "items": items,
            "technical_sections": technical_sections,
            "categories": list(categories.keys())
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False)
        
        total_time = time.time() - start_time
        print(f"Saved to {output_path}", flush=True)
        print(f"=== EXTRACTION COMPLETED in {total_time:.1f}s ===", flush=True)
        return True
        
    except Exception as e:
        print(f"FATAL ERROR: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Script started", flush=True)
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