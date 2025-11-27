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
        print("=== STARTING SEMANTIC EXTRACTION (Prices + Surrounding Technical Content) ===", flush=True)
        
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
        
        # Process ALL pages
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
        
        print("Building SEMANTIC extraction request...", flush=True)
        
        content = [
            {"type": "text", "text": """You are extracting data from a supplier quotation to transfer it into our company template.

CRITICAL: Extract COMPLETE information for each priced item by reading the ENTIRE document semantically - not just table cells.

═══════════════════════════════════════════════════════════════
EXTRACTION STRATEGY
═══════════════════════════════════════════════════════════════

1. FIND ALL PRICED ITEMS
   Scan for sections with pricing:
   - Main equipment tables
   - Accessories sections  
   - Format changes / spare parts
   - Packing / shipping
   - Any line with: price, "Included", "On request", "To be quoted"

2. FOR EACH ITEM, EXTRACT COMPLETE CONTEXT
   
   A) BASIC INFO (from table):
      - item_name: Short title from table
      - quantity: Number (default "1")
      - unit_price: Exact format ("€96.900,00" OR "Included" OR "On request")
      - total_price: Same format
      - category: Section name ("Main Equipment", "ACCESSORIES", etc.)
   
   B) FULL DESCRIPTION - READ THE ENTIRE DOCUMENT:
      Look for technical content NEAR this item:
      - Paragraphs ABOVE/BELOW the pricing table
      - Technical sections on FOLLOWING pages
      - Detailed product descriptions
      - How the equipment works
      - Components included
      - Materials and construction
      
      COMBINE INTO ONE DESCRIPTION (100-300 words):
      - Product overview
      - How it works / what it does
      - Key features and components
      - Materials and specifications
      - Standards compliance
      - Included accessories
      
      EXAMPLE for a distillation unit:
      "DISCONTINUOUS DISTILLATION UNIT C27 for wine, fermented grapes and other fruits - working with indirect steam at max. 0,5 bar, operating at atmospheric pressure. Alembic capacity: 1000 Litres. The unit features copper and stainless steel construction with a tall vertical column and spherical alembic. Complete with heating jacket, temperature control system, cooling condenser, and product collection vessel. The alembic operates by heating the wine or fruit must in the lower vessel, with vapors rising through the copper column for distillation. Cooling water circulates through the condenser to separate alcohol from water. All contact surfaces are food-grade stainless steel AISI 304. Includes safety pressure relief valve and temperature monitoring. The unit is made according to EC regulations 2006/42/EC, 2014/30/EU, 2014/35/EU."
   
   C) SPECIFICATIONS (from spec tables):
      Extract structured technical data:
      - Dimensions: "4000 x 1530 x 4500 mm"
      - Weight: "865 Kg"  
      - Capacity: "1000 Litres"
      - Power: "2,3 kW"
      - Pressure ratings: "Max 0.5 bar"
      - Materials: "AISI 304 stainless steel"
      - Standards: "CE certified, EC 2006/42"
      
      Format as single line: "Capacity: 1000L, Dimensions: 4000x1530x4500mm, Weight: 865kg, Power: 2.3kW, Max pressure: 0.5bar, Material: AISI 304, Standards: EC 2006/42/EC, 2014/30/EU, 2014/35/EU"

3. SEMANTIC READING RULES

   ✓ Read pages BEFORE pricing table (product introduction)
   ✓ Read pages AFTER pricing table (technical details)
   ✓ Look for equipment photos/diagrams - describe what you see
   ✓ Connect table items to surrounding text by name/model
   ✓ Include regulatory info (CE marks, standards)
   ✓ Preserve exact model numbers, part codes
   ✓ Keep technical terms in original language

4. PRICE STATES
   - Numeric: Keep EXACT format with dots/commas: "€96.900,00"
   - Included: "Included" (when price is 0 or says "included")
   - On request: "On request" (when "to be quoted", "can be offered")

═══════════════════════════════════════════════════════════════
OUTPUT FORMAT
═══════════════════════════════════════════════════════════════

Return ONLY valid JSON array:

[{
  "category": "Main Equipment",
  "item_name": "DISCONTINUOUS DISTILLATION UNIT C27",
  "quantity": "1",
  "unit_price": "€96.900,00",
  "total_price": "€96.900,00",
  "description": "DISCONTINUOUS DISTILLATION UNIT C27 for wine, fermented grapes and other fruits - working with indirect steam at max. 0,5 bar, operating at atmospheric pressure. Alembic capacity: 1000 Litres. The unit features copper and stainless steel construction with a tall vertical column and spherical alembic. Complete with heating jacket, temperature control system, cooling condenser, and product collection vessel. The alembic operates by heating the wine or fruit must in the lower vessel, with vapors rising through the copper column for distillation. Cooling water circulates through the condenser to separate alcohol from water. All contact surfaces are food-grade stainless steel AISI 304. Includes safety pressure relief valve and temperature monitoring. The unit is made according to EC regulations 2006/42/EC, 2014/30/EU, 2014/35/EU.",
  "specifications": "Capacity: 1000 Litres, Maximum steam pressure: 0.5 bar, Steam working pressure: 0.2-0.5 bar, Installed electric power: 2.3 kW, Dimensions: 4000 x 1530 x 4500 mm, Weight: 865 Kg, Material: Copper and stainless steel construction, Standards: EC 2006/42/EC, 2014/30/EU, 2014/35/EU",
  "details": "Custom Tariff 8419 4000. Other features as per enclosed description."
},
{
  "category": "Packing",
  "item_name": "Packing (wooden crates) and loading on truck",
  "quantity": "1",
  "unit_price": "€1.830,00",
  "total_price": "€1.830,00",
  "description": "Packing in wooden crates and loading on truck for safe transportation. Professional export packaging with wood crating to protect equipment during international shipping.",
  "specifications": "",
  "details": ""
}]

CRITICAL REMINDERS:
- Read ENTIRE document semantically - not just table cells
- Build descriptions from ALL content about each item (100-300 words)
- Look for technical details on pages AFTER the pricing table
- Extract complete specifications from spec tables
- Preserve exact pricing format
- Return ONLY valid JSON - no markdown, no explanations
"""}
        ]
        
        for img_data in image_data_list:
            content.append({"type": "image_url", "image_url": {"url": img_data}})
        
        print("Calling OpenAI Vision API with SEMANTIC extraction...", flush=True)
        print("Token limit: 8000", flush=True)
        
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
        items = json.loads(extracted_json)
        
        print(f"✓ Extracted {len(items)} items", flush=True)
        
        if len(items) == 0:
            print("ERROR: No items extracted", flush=True)
            return False
        
        # Quality metrics
        items_with_description = sum(1 for item in items if len(item.get('description', '')) > 100)
        items_with_specs = sum(1 for item in items if item.get('specifications', '').strip())
        
        total_desc_length = sum(len(item.get('description', '')) for item in items)
        avg_desc_length = total_desc_length // len(items) if items else 0
        
        print("\n" + "="*60, flush=True)
        print("SEMANTIC EXTRACTION QUALITY", flush=True)
        print("="*60, flush=True)
        print(f"Total items: {len(items)}", flush=True)
        print(f"Items with full descriptions (>100 chars): {items_with_description} ({items_with_description*100//len(items)}%)", flush=True)
        print(f"Items with specifications: {items_with_specs} ({items_with_specs*100//len(items)}%)", flush=True)
        print(f"Average description length: {avg_desc_length} characters", flush=True)
        print("="*60 + "\n", flush=True)
        
        # Show sample
        if items:
            sample = items[0]
            print("SAMPLE EXTRACTION:", flush=True)
            print(f"  Item: {sample.get('item_name', 'N/A')}", flush=True)
            print(f"  Description: {sample.get('description', '')[:200]}...", flush=True)
            print(f"  Description length: {len(sample.get('description', ''))} chars", flush=True)
            print(f"  Specifications: {sample.get('specifications', 'N/A')[:100]}...", flush=True)
            print("", flush=True)
        
        # Group by category
        categories = {}
        for item in items:
            cat = item.get("category", "Main Items")
            if cat not in categories:
                categories[cat] = []
            categories[cat].append(item)
        
        print(f"Grouped into {len(categories)} categories:", flush=True)
        for cat, cat_items in categories.items():
            print(f"  - {cat}: {len(cat_items)} items", flush=True)
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # Save
        output_data = {
            "items": items,
            "categories": list(categories.keys()),
            "extraction_version": "SV13_Semantic"
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False)
        
        print(f"✓ Saved to {output_path}", flush=True)
        print("=== SEMANTIC EXTRACTION COMPLETED ===", flush=True)
        return True
        
    except Exception as e:
        print(f"FATAL ERROR: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Semantic Extraction Script Started", flush=True)
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