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
        print("=== STARTING ENHANCED EXTRACTION (SV12) ===", flush=True)
        
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
        
        print("Building enhanced OpenAI request...", flush=True)
        
        content = [
            {"type": "text", "text": """Extract EVERY item with a price or marked "Included" from this quotation with COMPLETE technical information.

CRITICAL EXTRACTION RULES:
═══════════════════════════════════════════════════════════════

1. SCAN ALL PAGES - Extract from entire document until "GENERAL SALE TERMS" or end

2. EXTRACT FROM ALL SECTIONS:
   • Main equipment/machinery
   • Economic Offer table (main pricing table)
   • Format Changes
   • Accessories sections
   • Further Options
   • Packing options
   • Any other itemized sections

3. FOR EACH ITEM, EXTRACT COMPLETE INFORMATION:

   A) BASIC INFO (same as before):
      - category: Section name (e.g., "Main Equipment", "FORMAT CHANGES", "ACCESSORIES")
      - item_name: Short item title
      - quantity: Number (default "1" if not shown)
      - unit_price: Price (numeric like "€324.400,00" OR "Included" OR "On request")
      - total_price: Total (same format as unit_price)

   B) COMPLETE DESCRIPTION (NEW - CRITICAL):
      - description: FULL technical description (50-200 words)
        * Extract COMPLETE text from document, NEVER truncate
        * Include ALL details, specifications, features
        * Preserve model numbers, part codes exactly
        * Keep technical terms in original language
        * Include parenthetical details like "(included)", "(optional)"
        * Maintain bullet points if present
        * Continue until item ends or next item starts
        * NEVER write "..." or "[more details]" - extract EVERYTHING

   C) TECHNICAL SPECIFICATIONS (NEW):
      - specifications: Structured technical details
        * Dimensions (e.g., "ø15mm", "300x200mm")
        * Capacity (e.g., "0.33L", "5000 bph")
        * Materials (e.g., "AISI 304", "stainless steel")
        * Power (e.g., "3kW", "220V")
        * Weight, speed, temperature ranges
        * Standards (e.g., "CE certified", "ISO compliant")
        * If none found, use empty string ""

   D) IMAGE INFORMATION (NEW):
      - has_image: true/false - Is there an image/photo/diagram near this item?
      - image_description: If has_image=true, describe what the image shows:
        * Equipment appearance, shape, color
        * Key visible components
        * Installation context if shown
        * If no image, use empty string ""

   E) ADDITIONAL DETAILS:
      - details: Any other notes, installation info, certifications not in specifications

4. PRICE STATES (UNCHANGED):
   • Numeric: Keep exact format "€324.400,00" (preserve dots/commas)
   • Included: "Included" (when text says "Included" or price is 0)
   • On request: "On request" (when "Can be offered", "To be quoted", "Please inquire")

5. EXTRACTION QUALITY REQUIREMENTS:
   ✅ Extract COMPLETE descriptions - no truncation
   ✅ Preserve ALL technical specifications
   ✅ Detect ALL images in document
   ✅ Keep model numbers and part codes exact
   ✅ Maintain formatting (bullets, lists)
   ✅ Continue until explicit section end

RETURN ONLY JSON ARRAY:

[{
  "category": "Main Equipment",
  "item_name": "CAN ISO 20/2 S - clock wisely running direction",
  "quantity": "1",
  "unit_price": "€324.400,00",
  "total_price": "€324.400,00",
  "description": "Complete can filling machine based on one size of 0.33L standard aluminium can. The system includes Rolls kit for 1st and 2nd operation with chuck in stainless steel AISI 304. Features electronic filling level control, automatic can handling system with star wheels, and integrated CIP cleaning system. Designed for beverage industry applications with sanitary design meeting FDA standards. Machine includes operator touch-screen panel (7-inch color display), safety guards, and emergency stop systems compliant with CE regulations.",
  "specifications": "Capacity: 2000 cph, Can size: 0.33L (ø66mm), Material: AISI 304 stainless steel, Power: 3kW 220V, Dimensions: 1800x1200x2100mm, Weight: 450kg, Operating pressure: 2-4 bar",
  "has_image": true,
  "image_description": "Photo shows the complete filling machine in stainless steel construction with control panel on right side, filling heads visible at top, and conveyor system at base. Machine has professional industrial appearance with safety guards.",
  "details": "Installation includes mechanical setup, electrical connection, and initial operator training. 12-month warranty included."
},
{
  "category": "ACCESSORIES",
  "item_name": "FEEDING PUMP",
  "quantity": "1",
  "unit_price": "On request",
  "total_price": "On request",
  "description": "Centrifugal feeding pump designed for beverage products. Can be offered and configured according to the specific product characteristics being filled. Pump features sanitary design with stainless steel construction, CIP compatible components, and FDA-approved seals. Includes variable frequency drive for flow rate adjustment and pressure monitoring system.",
  "specifications": "Material: AISI 316L stainless steel, Flow rate: 50-200 L/min (adjustable), Power: 2.2kW, Connection: 2-inch tri-clamp",
  "has_image": false,
  "image_description": "",
  "details": "Pump selection depends on product viscosity and desired flow rate. Technical consultation available."
}]

CRITICAL REMINDERS:
• Extract FULL descriptions (50-200 words each) - NEVER truncate
• Include ALL specifications you can find
• Mark images accurately (has_image true/false)
• Continue extracting until document ends or "GENERAL SALE TERMS"
• Return ONLY valid JSON - no explanations
"""}
        ]
        
        for img_data in image_data_list:
            content.append({"type": "image_url", "image_url": {"url": img_data}})
        
        print("Calling OpenAI Vision API with enhanced extraction...", flush=True)
        print("Token limit: 8000 (increased for full descriptions)", flush=True)
        
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": content}],
            max_tokens=8000,  # INCREASED from 6000 to 8000
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
        
        # Calculate quality metrics
        items_with_description = sum(1 for item in items if item.get('description', '').strip() and len(item.get('description', '')) > 50)
        items_with_specs = sum(1 for item in items if item.get('specifications', '').strip())
        items_with_images = sum(1 for item in items if item.get('has_image', False))
        
        total_desc_length = sum(len(item.get('description', '')) for item in items)
        avg_desc_length = total_desc_length // len(items) if items else 0
        
        print("\n" + "="*60, flush=True)
        print("EXTRACTION QUALITY METRICS (SV12)", flush=True)
        print("="*60, flush=True)
        print(f"Total items extracted: {len(items)}", flush=True)
        print(f"Items with full descriptions (>50 chars): {items_with_description} ({items_with_description*100//len(items)}%)", flush=True)
        print(f"Items with specifications: {items_with_specs} ({items_with_specs*100//len(items)}%)", flush=True)
        print(f"Items with images detected: {items_with_images} ({items_with_images*100//len(items)}%)", flush=True)
        print(f"Average description length: {avg_desc_length} characters", flush=True)
        print("="*60 + "\n", flush=True)
        
        # Show sample extraction
        if items:
            sample = items[0]
            print("SAMPLE EXTRACTION (First Item):", flush=True)
            print(f"  Item: {sample.get('item_name', 'N/A')[:60]}...", flush=True)
            print(f"  Description length: {len(sample.get('description', ''))} chars", flush=True)
            print(f"  Has specifications: {'Yes' if sample.get('specifications', '').strip() else 'No'}", flush=True)
            print(f"  Has image: {'Yes' if sample.get('has_image', False) else 'No'}", flush=True)
            print("", flush=True)
        
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
        
        # Save with quality metrics
        output_data = {
            "items": items,
            "categories": list(categories.keys()),
            "extraction_version": "SV12_Enhanced",
            "quality_metrics": {
                "total_items": len(items),
                "items_with_descriptions": items_with_description,
                "items_with_specifications": items_with_specs,
                "items_with_images": items_with_images,
                "average_description_length": avg_desc_length,
                "description_coverage_percent": items_with_description*100//len(items),
                "specifications_coverage_percent": items_with_specs*100//len(items),
                "image_detection_percent": items_with_images*100//len(items)
            }
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False)
        
        print(f"✓ Saved to {output_path}", flush=True)
        print("=== ENHANCED EXTRACTION COMPLETED (SV12) ===", flush=True)
        return True
        
    except Exception as e:
        print(f"FATAL ERROR: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Enhanced Extraction Script Started (SV12)", flush=True)
    pdf_path = os.path.join(UPLOAD_FOLDER, "offer1.pdf")
    output_path = os.path.join(OUTPUT_FOLDER, "items_offer1.json")
    
    if not os.path.exists(pdf_path):
        print("ERROR: PDF not found at " + pdf_path, flush=True)
        sys.exit(1)
    
    success = extract_items_from_pdf(pdf_path, output_path)
    
    if not success:
        print("Enhanced extraction failed", flush=True)
        sys.exit(1)
    
    print("COMPLETED SUCCESSFULLY (SV12)", flush=True)
    sys.exit(0)