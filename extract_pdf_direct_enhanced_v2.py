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

# PROMPT 1: Extract ALL data from Offer 1 (Supplier Quotation)
EXTRACTION_PROMPT = """PROMPT 1: EXTRACT ALL DATA FROM OFFER 1 (SUPPLIER QUOTATION)
==============================================================

You are an experienced procurement specialist who manually rewrites supplier quotations into company-branded offers. This is what you do every day at work.

YOUR DAILY TASK:
A supplier sent you their quotation (Offer 1). Your job is to extract ALL information from it and create a clean, professional offer for your client using your company's branded template (Offer 2).

WHAT YOU EXTRACT FROM SUPPLIER QUOTATION (Offer 1):

1. PRICING DATA (Every Item):
   - Item name/description with model numbers
   - Technical specifications
   - Quantity
   - Unit price with currency (€, $, £, etc.)
   - Total price
   - Notes (Included, Optional, On request, etc.)

2. TECHNICAL CONTENT:
   - Product descriptions (paragraphs of text)
   - Technical specifications (feature lists, bullet points)
   - Specification tables (dimensions, capacity, power, etc.)
   - Installation requirements
   - Warranty information
   - Compliance certifications

3. VISUAL ELEMENTS:
   - Product photos/images
   - Technical diagrams
   - Logos (if any)
   - Charts or infographics

HOW YOU WORK (Like a Real Human):

1. READ THE ENTIRE DOCUMENT
   - Start from page 1, read every section
   - Scan all pages until you see "Terms and Conditions", "Payment Terms", or document end
   - Check for continuation indicators ("See next page", "Continued", etc.)

2. EXTRACT FROM ALL SECTIONS
   - Main equipment/products table
   - Optional accessories
   - Add-ons and upgrades
   - Packing options
   - Format changes
   - Additional services
   - ANY section that has pricing or technical details

3. CAPTURE TECHNICAL DESCRIPTIONS
   - Copy full product descriptions (not just item names)
   - Extract feature lists and specifications
   - Note any special requirements or conditions
   - Identify which images belong to which products

4. HANDLE DIFFERENT PRICE FORMATS
   - Numeric prices: €324.400,00 or $15,400.00 (preserve exact format)
   - Included items: "Included" (when price is 0 or marked as included)
   - Quote on request: "On request" (when it says "Can be offered", "To be quoted", "Please inquire")

5. MULTI-LANGUAGE RECOGNITION
   - English, Ukrainian, Russian, Italian, German, French, Spanish
   - Recognize column headers in any language
   - Preserve original language in descriptions

VALIDATION BEFORE FINISHING (Self-Check):
☑ Did I read the entire document to the end?
☑ Did I extract prices from ALL sections (not just main table)?
☑ Did I capture technical descriptions and specifications?
☑ Did I identify all images and which products they relate to?
☑ Did I check for optional/accessory sections?
☑ Are there any price indicators after my last item? (If yes → go back)
☑ Did I preserve the exact price format with thousand separators?

RETURN FORMAT - Complete JSON with all extracted data:

{
  "items": [
    {
      "category": "Main Equipment",
      "item_name": "CAN FILLER ISO 20/2 S",
      "technical_description": "Full paragraph describing features, capabilities, included components...",
      "specifications": {
        "model": "ISO 20/2 S",
        "capacity": "20 filling valves, 2 seaming heads",
        "can_size": "0.33L standard aluminum",
        "direction": "Clockwise"
      },
      "quantity": "1",
      "unit_price": "€324.400,00",
      "total_price": "€324.400,00",
      "notes": "",
      "related_images": ["image_1", "image_2"]
    }
  ],
  "technical_sections": [
    {
      "title": "Features of the rinsing turret",
      "content": "Full text content of this section...",
      "type": "text_paragraph"
    },
    {
      "title": "Technical Specifications",
      "content": ["Feature 1", "Feature 2", "Feature 3"],
      "type": "bullet_list"
    }
  ],
  "images": [
    {
      "id": "image_1",
      "description": "Can filler machine front view",
      "related_items": ["CAN FILLER ISO 20/2 S"]
    }
  ],
  "document_metadata": {
    "currency": "EUR",
    "total_items": 47,
    "has_optional_sections": true,
    "languages_detected": ["English", "Italian"]
  }
}

Now, analyze this supplier quotation and extract everything as if you were manually rewriting it:"""

def extract_items_from_pdf(pdf_path, output_path):
    try:
        print("=== STARTING FLEXIBLE EXTRACTION (Using PROMPT 1) ===", flush=True)
        
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
        
        print("Building request with PROMPT 1...", flush=True)
        
        # Use PROMPT 1 approach
        content = [{"type": "text", "text": EXTRACTION_PROMPT}]
        
        for img_data in image_data_list:
            content.append({"type": "image_url", "image_url": {"url": img_data}})
        
        print("Calling OpenAI Vision API with PROMPT 1...", flush=True)
        
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": content}],
            max_tokens=8000,
            temperature=0
        )
        
        print("Received response from OpenAI", flush=True)
        
        extracted_json = response.choices[0].message.content.strip()
        
        # Clean JSON formatting
        if extracted_json.startswith("```json"):
            extracted_json = extracted_json.replace("```json", "").replace("```", "").strip()
        elif extracted_json.startswith("```"):
            extracted_json = extracted_json.replace("```", "").strip()
        
        print("Parsing JSON...", flush=True)
        full_data = json.loads(extracted_json)
        
        items = full_data.get("items", [])
        technical_sections = full_data.get("technical_sections", [])
        images = full_data.get("images", [])
        metadata = full_data.get("document_metadata", {})
        
        print(f"✓ Extracted {len(items)} items", flush=True)
        print(f"✓ Extracted {len(technical_sections)} technical sections", flush=True)
        print(f"✓ Found {len(images)} images", flush=True)
        print(f"✓ Metadata: {metadata}", flush=True)
        
        if len(items) == 0:
            print("ERROR: No items extracted", flush=True)
            return False
        
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
        
        # Save full data with flexible structure
        output_data = {
            "extraction_method": "PROMPT_1_Flexible",
            "items": items,
            "technical_sections": technical_sections,
            "images": images,
            "document_metadata": metadata,
            "categories": list(categories.keys())
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False)
        
        print(f"✓ Saved to {output_path}", flush=True)
        print("=== FLEXIBLE EXTRACTION COMPLETED ===", flush=True)
        return True
        
    except Exception as e:
        print(f"FATAL ERROR: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Flexible Extraction Script Started (PROMPT 1)", flush=True)
    pdf_path = os.path.join(UPLOAD_FOLDER, "offer1.pdf")
    output_path = os.path.join(OUTPUT_FOLDER, "items_offer1.json")
    
    if not os.path.exists(pdf_path):
        print("ERROR: PDF not found at " + pdf_path, flush=True)
        sys.exit(1)
    
    success = extract_items_from_pdf(pdf_path, output_path)
    
    if not success:
        print("Flexible extraction failed", flush=True)
        sys.exit(1)
    
    print("COMPLETED SUCCESSFULLY (PROMPT 1)", flush=True)
    sys.exit(0)