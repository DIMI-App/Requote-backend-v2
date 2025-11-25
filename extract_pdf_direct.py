import os
import sys
import json
import openai
import fitz
import base64

openai.api_key = os.environ.get('OPENAI_API_KEY')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def extract_items_from_pdf(pdf_path, output_path):
    try:
        print("=== STARTING ENHANCED EXTRACTION (Prices + Technical Descriptions) ===", flush=True)
        
        if not openai.api_key:
            print("ERROR: OPENAI_API_KEY not set", flush=True)
            return False
        
        print("OpenAI key found", flush=True)
        print("Reading PDF: " + pdf_path, flush=True)
        
        if not os.path.exists(pdf_path):
            print("ERROR: PDF file not found", flush=True)
            return False
        
        print("Opening PDF with PyMuPDF...", flush=True)
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        print(f"PDF has {total_pages} pages", flush=True)
        
        # Process ALL pages (up to 15 for typical quotes)
        max_pages = min(15, total_pages)
        print(f"Processing first {max_pages} pages", flush=True)
        
        image_data_list = []
        for page_num in range(max_pages):
            print(f"Converting page {page_num + 1}...", flush=True)
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            img_bytes = pix.tobytes("png")
            img_base64 = base64.b64encode(img_bytes).decode('utf-8')
            image_data_list.append(f"data:image/png;base64,{img_base64}")
            print(f"Page {page_num + 1}: converted ({len(img_base64)} bytes)", flush=True)
        
        doc.close()
        print("All pages converted", flush=True)
        
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
   - Installation requirements
   - Material specifications
   - Performance characteristics

For each technical section:
   - section_title: Title or heading of the section
   - content_type: "paragraph" | "bullet_list" | "spec_table" | "features"
   - content: Full text content
   - page_location: "before_price_table" | "after_price_table" | "between_items"

IMPORTANT:
- Extract technical content that appears OUTSIDE of pricing tables
- Include product descriptions, feature lists, specifications
- Preserve paragraph structure and formatting
- Include bullet points as they appear
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
    },
    {
      "section_title": "Technical Specifications",
      "content_type": "spec_table",
      "content": "Production capacity: 12,000 bph\\nPower: 15 kW\\nDimensions: 2500x1800x2200 mm",
      "page_location": "after_price_table"
    }
  ]
}

PRICE STATE EXAMPLES:
- "€324.400,00" → unit_price: "€324.400,00"
- "Included" → unit_price: "Included"
- "Can be offered" → unit_price: "On request"
"""}
        ]
        
        for img_data in image_data_list:
            content.append({"type": "image_url", "image_url": {"url": img_data}})
        
        print("Calling OpenAI Vision API...", flush=True)
        
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": content}],
            max_tokens=8000,
            temperature=0
        )
        
        print("Received response from OpenAI", flush=True)
        
        extracted_json = response.choices[0].message.content.strip()
        
        if extracted_json.startswith("```json"):
            extracted_json = extracted_json.replace("```json", "").replace("```", "").strip()
        elif extracted_json.startswith("```"):
            extracted_json = extracted_json.replace("```", "").strip()
        
        print("Parsing JSON...", flush=True)
        full_data = json.loads(extracted_json)
        
        items = full_data.get("items", [])
        technical_sections = full_data.get("technical_sections", [])
        
        print(f"Extracted {len(items)} items", flush=True)
        print(f"Extracted {len(technical_sections)} technical sections", flush=True)
        
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
        
        # Save full data with both items and technical sections
        output_data = {
            "items": items,
            "technical_sections": technical_sections,
            "categories": list(categories.keys())
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False)
        
        print(f"Saved to {output_path}", flush=True)
        print("=== EXTRACTION COMPLETED ===", flush=True)
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