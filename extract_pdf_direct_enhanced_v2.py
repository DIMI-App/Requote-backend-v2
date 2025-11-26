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
EXTRACTION_PROMPT = """PROMPT 1: EXTRACT EVERY LINE ITEM FROM SUPPLIER QUOTATION
================================================================

You are extracting data from a supplier quotation. Your job is to extract EVERY SINGLE LINE that has a price.

CRITICAL EXTRACTION RULES:

1. **EXTRACT EVERY LINE ITEM SEPARATELY**
   - Main equipment (machines, systems, units)
   - Packing/crating (wooden crates, packaging)
   - Shipping/loading (loading on truck, freight)
   - Installation services
   - Optional accessories
   - Training
   - ANY line that shows a price or cost

2. **EACH LINE = SEPARATE ITEM**
   Example from quotation:
   ```
   1  DISTILLATION UNIT C27    €96,900.00
      Packing (wooden crates)   €1,830.00
   ```
   This is **2 ITEMS**, not 1:
   - Item 1: DISTILLATION UNIT C27 - €96,900.00
   - Item 2: Packing (wooden crates) - €1,830.00

3. **SCAN ENTIRE DOCUMENT**
   - Read ALL pages from start to finish
   - Stop at "Terms and Conditions", "Payment Terms", or "Exclusions"
   - Extract from tables AND from text paragraphs with prices

4. **EXTRACT TECHNICAL CONTENT**
   - Full product descriptions (multi-paragraph text blocks)
   - Component lists (bullet points describing parts)
   - Technical specification tables (capacity, dimensions, power, etc.)
   - Feature descriptions
   - Images and diagrams

5. **HANDLE PRICES**
   - Numeric: "€96,900.00" or "$15,400" (preserve exact format)
   - Included: "Included" (when marked as included or price is 0)
   - On request: "On request" (when "to be quoted", "on demand", etc.)

**VALIDATION CHECKLIST** (Ask yourself before returning):
☑ Did I extract EVERY price line? (main items + packing + options)
☑ Did I treat packing/crating as SEPARATE items?
☑ Did I scan the ENTIRE document to the end?
☑ Did I extract technical descriptions AND specification tables?
☑ Did I identify which images belong to which products?

**RETURN FORMAT:**

{
  "items": [
    {
      "category": "Main Equipment",
      "item_name": "DISCONTINUOUS DISTILLATION UNIT C27",
      "technical_description": "For wine, fermented grapes and other fruits - working with indirect steam at max. 0,5 bar, operating at atmospheric pressure. Alembic capacity: 1000 Litres",
      "specifications": {
        "model": "C27",
        "capacity": "1000 Litres",
        "working_pressure": "0,5 bar"
      },
      "quantity": "1",
      "unit_price": "€96.900,00",
      "total_price": "€96.900,00",
      "notes": "Custom Tariff 8419 4000",
      "related_images": ["C27_image"]
    },
    {
      "category": "Packing",
      "item_name": "Packing (wooden crates) and loading on truck",
      "technical_description": "",
      "specifications": {},
      "quantity": "1",
      "unit_price": "€1.830,00",
      "total_price": "€1.830,00",
      "notes": "",
      "related_images": []
    }
  ],
  "technical_sections": [
    {
      "title": "Features of the Discontinuous Distillation Unit C27",
      "content": "An alembic in stainless steel and copper with a working capacity of 10 hl, equipped with a slow revolving stirrer...",
      "type": "text_paragraph"
    },
    {
      "title": "C27 Technical Data",
      "content": {
        "table_headers": ["NAME", "UNIT", "VALUE"],
        "table_rows": [
          ["Raw material production capacity per cycle", "litres", "1.000"],
          ["Total time each cycle", "hours", "2.5 / 4.5"]
        ]
      },
      "type": "specification_table"
    }
  ],
  "images": [
    {
      "id": "C27_image",
      "description": "C27 distillation unit equipment photo",
      "related_items": ["DISCONTINUOUS DISTILLATION UNIT C27"]
    }
  ],
  "document_metadata": {
    "currency": "EUR",
    "total_items": 4,
    "has_optional_sections": false,
    "languages_detected": ["English"]
  }
}

**CRITICAL REMINDERS:**
- Packing is a SEPARATE item, not part of equipment
- Extract specification TABLES as tables (with headers and rows)
- Each price line = one item in the array
- Continue scanning until you see "EXCLUSIONS" or "Terms and Conditions"

Now extract EVERYTHING from this quotation:"""

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