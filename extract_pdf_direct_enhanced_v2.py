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

# NEW PROMPT 1: SEMANTIC EXTRACTION FROM SUPPLIER QUOTATION
EXTRACTION_PROMPT = """PROMPT 1: SEMANTIC EXTRACTION FROM SUPPLIER QUOTATION (OFFER 1)

================================================================================
CONTEXT UNDERSTANDING FIRST
================================================================================

Read this supplier quotation carefully and answer these questions first:

1. INDUSTRY & PRODUCT CATEGORY
   What is this offer about?
   Examples:
   - "Wine production equipment - fermentation tanks"
   - "Bottling machinery - capsuling machines"
   - "Industrial pumps - centrifugal pumps"
   - "Distillation equipment - alembic stills"

2. MAIN EQUIPMENT IDENTIFICATION
   What is the PRIMARY equipment being offered?
   - Look for the most prominent/expensive item
   - Look for what everything else relates to
   - It might not have a label like "Main equipment"
   - There may be multiple main equipment items
   
   Examples:
   - "Corking machine SVP 456" = main equipment
   - "Tank 5000L + Tank 3000L" = two main equipment items

3. ADDITIONAL ITEMS
   What else is included?
   - Accessories (complete the main equipment)
   - Options (customer can choose to add)
   - Services (installation, training, warranty)
   - Spare parts
   - Packing/shipping

4. OFFER STRUCTURE
   How is this offer organized?
   - Single product with options?
   - Multiple independent machines?
   - Complete production line?
   - Equipment + installation package?

Provide your understanding in simple language first.

================================================================================
IDENTIFY SEMANTIC SECTIONS
================================================================================

Now identify WHERE different types of information are located:

LOOK FOR THESE PURPOSES (not labels):

📌 MAIN EQUIPMENT NAME
   - Where is it? (title, first table cell, header, bold text)
   - Recognition: Most prominent equipment name
   - Note: "Corking machine SVP 456" = equipment name (even without label)

📌 PRICING INFORMATION
   - Where are prices? (table, list, inline text)
   - Recognition: Currency symbols, numbers with decimals
   - Note: A table with numbers and € = pricing (even if no header says "prices")

📌 TECHNICAL SPECIFICATIONS
   - Where are specs? (table, bullet points, paragraph)
   - Recognition: Parameters with values and units
   - Note: "Capacity: 1000L, Material: Steel" = technical specs (even if no header)

📌 TECHNICAL DESCRIPTION
   - Where is narrative text? (paragraphs, bullets)
   - Recognition: Explains what equipment does, how it works
   - Note: This is different from specs - it's descriptive text

📌 COMMERCIAL TERMS
   - Where are business terms? (payment, delivery, warranty)
   - Recognition: Usually at end of document

📌 IMAGES/PHOTOS
   - Where are visuals? (equipment photos, drawings)
   - Note their locations and what they show

================================================================================
EXTRACT WITH SEMANTIC STRUCTURE
================================================================================

Based on your understanding, extract all information in this JSON format:

{
  "context_understanding": {
    "industry": "wine production equipment",
    "main_product_category": "fermentation tanks",
    "offer_type": "single_equipment_with_accessories"
  },
  
  "main_equipment": {
    "name": "Stainless Steel Fermentation Tank 5000L",
    "identified_from": "Page 1, bold title at top",
    "reasoning": "This is the main product - highest price, entire document focuses on it",
    "category": "fermentation_tank",
    "note": "If multiple main equipment items exist, list them all"
  },
  
  "pricing_items": [
    {
      "type": "main_equipment",
      "description": "Full description as written in offer",
      "identified_as_main_because": "Highest price and all other items relate to it",
      "quantity": 1,
      "unit": "pcs",
      "unit_price": 45000,
      "total": 45000,
      "currency": "EUR"
    },
    {
      "type": "accessory",
      "description": "Temperature control system",
      "identified_as_accessory_because": "Completes the tank functionality",
      "quantity": 1,
      "unit": "pcs",
      "unit_price": 3500,
      "total": 3500,
      "currency": "EUR"
    },
    {
      "type": "packing",
      "description": "Wooden crate packing and loading",
      "identified_as_packing_because": "Transport/packaging service",
      "quantity": 1,
      "unit": "lot",
      "unit_price": 1200,
      "total": 1200,
      "currency": "EUR"
    },
    {
      "type": "option",
      "description": "Additional insulation layer",
      "identified_as_option_because": "Customer can choose to add or not",
      "quantity": 1,
      "unit": "pcs",
      "unit_price": 800,
      "total": 800,
      "currency": "EUR",
      "is_optional": true
    }
  ],
  
  "technical_specifications": [
    {
      "parameter": "Capacity",
      "value": "5000",
      "unit": "liters",
      "equipment": "main_tank"
    },
    {
      "parameter": "Material",
      "value": "AISI 304 stainless steel",
      "unit": null,
      "equipment": "main_tank"
    },
    {
      "parameter": "Working pressure",
      "value": "3",
      "unit": "bar",
      "equipment": "main_tank"
    },
    {
      "parameter": "Temperature range",
      "value": "-5 to +80",
      "unit": "°C",
      "equipment": "main_tank"
    }
  ],
  
  "technical_description": {
    "full_text": "Complete paragraph(s) describing equipment functionality, construction, features, how it works. This should be 50-500 words of narrative text.",
    "identified_from": "Page 2, section titled 'Description' or 'Technical Information'",
    "reasoning": "This section explains how the equipment works, not just specifications",
    "key_features": [
      "Double-wall construction with insulation",
      "Food-grade stainless steel",
      "Integrated cooling system",
      "Easy-access inspection ports"
    ]
  },
  
  "images": [
    {
      "type": "equipment_photo",
      "description": "Main tank front view",
      "location": "Page 3",
      "related_equipment": "main_tank"
    },
    {
      "type": "technical_drawing",
      "description": "Dimensions diagram",
      "location": "Page 4",
      "related_equipment": "main_tank"
    }
  ],
  
  "commercial_terms": {
    "payment_terms": "30% advance, 70% before delivery",
    "delivery_time": "10-12 weeks from order",
    "delivery_terms": "EXW factory",
    "warranty": "24 months",
    "validity": "30 days",
    "transport": "Excluded",
    "packing": "Included in offer",
    "starting_up": "Excluded - quoted separately",
    "voltage_connections": "Standard connections",
    "exclusions_summary": "List what is NOT included in the offer"
  },
  
  "pricing_summary": {
    "subtotal": 49700,
    "vat_rate": 20,
    "vat_amount": 9940,
    "total": 59640,
    "currency": "EUR"
  },
  
  "certifications": [
    "CE certified",
    "ISO 9001:2015",
    "Food-grade compliance"
  ]
}

================================================================================
CRITICAL INSTRUCTIONS
================================================================================

✅ DO:
- Use REASONING to explain WHY you identified something
- Focus on PURPOSE not labels
- Extract EVERYTHING - all prices, specs, descriptions
- Include "identified_from" to show where you found information
- If multiple main equipment items exist, list them all
- Preserve exact descriptions from document
- Note if information is missing (use null)

❌ DON'T:
- Don't rely only on section labels
- Don't invent data that isn't in document
- Don't skip items because they seem minor
- Don't combine items unless they're actually combined in source
- Don't assume - if unsure, explain your reasoning

================================================================================
EDGE CASES
================================================================================

IF multiple main equipment items:
- List each separately in pricing_items
- Create specs for each if they differ
- Note in main_equipment.note that there are multiple items

IF no clear main equipment:
- Choose the highest-priced or most prominent item
- Explain your reasoning clearly

IF prices are ranges (e.g., "€10,000 - €15,000"):
- Use middle value or note the range in description

IF specifications are embedded in description:
- Extract them to technical_specifications array
- Keep them in description too

IF document has multiple languages:
- Extract in original language
- Note which language in reasoning

Now extract EVERYTHING from this quotation:"""

def extract_items_from_pdf(pdf_path, output_path):
    try:
        print("=== STARTING SEMANTIC EXTRACTION (NEW PROMPT 1) ===", flush=True)
        
        if not openai.api_key:
            print("ERROR: OPENAI_API_KEY not set", flush=True)
            return False
        
        print(f"Reading PDF: {pdf_path}", flush=True)
        
        if not os.path.exists(pdf_path):
            print("ERROR: PDF file not found", flush=True)
            return False
        
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        print(f"PDF has {total_pages} pages", flush=True)
        
        # Convert pages to images
        max_pages = min(15, total_pages)
        image_data_list = []
        
        for page_num in range(max_pages):
            print(f"Converting page {page_num + 1}...", flush=True)
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_bytes = pix.tobytes("png")
            img_base64 = base64.b64encode(img_bytes).decode('utf-8')
            image_data_list.append(f"data:image/png;base64,{img_base64}")
        
        doc.close()
        
        # Build Vision API request
        content = [{"type": "text", "text": EXTRACTION_PROMPT}]
        for img_data in image_data_list:
            content.append({"type": "image_url", "image_url": {"url": img_data, "detail": "high"}})
        
        print("Calling GPT-4o Vision with NEW PROMPT 1...", flush=True)
        
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": content}],
            max_tokens=16000,
            temperature=0.1
        )
        
        extracted_json = response.choices[0].message.content.strip()
        
        # Clean JSON
        if extracted_json.startswith("```json"):
            extracted_json = extracted_json.replace("```json", "").replace("```", "").strip()
        elif extracted_json.startswith("```"):
            extracted_json = extracted_json.replace("```", "").strip()
        
        print("Parsing extraction...", flush=True)
        full_data = json.loads(extracted_json)
        
        context = full_data.get("context_understanding", {})
        pricing_items = full_data.get("pricing_items", [])
        tech_specs = full_data.get("technical_specifications", [])
        tech_desc = full_data.get("technical_description", {})
        
        print("=" * 60, flush=True)
        print("✅ SEMANTIC EXTRACTION COMPLETE (NEW PROMPT 1)", flush=True)
        print(f"   Industry: {context.get('industry', 'N/A')}", flush=True)
        print(f"   Category: {context.get('main_product_category', 'N/A')}", flush=True)
        print(f"   Pricing items: {len(pricing_items)}", flush=True)
        print(f"   Technical specs: {len(tech_specs)}", flush=True)
        
        # Convert to items format for compatibility with existing system
        items = []
        for pricing_item in pricing_items:
            item = {
                "category": pricing_item.get("type", "Main Equipment").replace("_", " ").title(),
                "item_name": pricing_item.get("description", ""),
                "quantity": str(pricing_item.get("quantity", 1)),
                "unit_price": f"{pricing_item.get('currency', 'EUR')} {pricing_item.get('unit_price', 0)}",
                "total_price": f"{pricing_item.get('currency', 'EUR')} {pricing_item.get('total', 0)}",
                "technical_description": "",
                "specifications": {},
                "notes": "",
                "related_images": []
            }
            items.append(item)
        
        # Save with NEW structure
        output_data = {
            "extraction_method": "NEW_PROMPT_1_Semantic",
            "context_understanding": context,
            "main_equipment": full_data.get("main_equipment", {}),
            "pricing_items": pricing_items,
            "technical_specifications": tech_specs,
            "technical_description": tech_desc,
            "images": full_data.get("images", []),
            "commercial_terms": full_data.get("commercial_terms", {}),
            "pricing_summary": full_data.get("pricing_summary", {}),
            "certifications": full_data.get("certifications", []),
            "items": items  # For compatibility
        }
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False)
        
        print(f"   Saved: {output_path}", flush=True)
        print("=" * 60, flush=True)
        
        return True
        
    except Exception as e:
        print(f"ERROR: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    pdf_path = os.path.join(UPLOAD_FOLDER, "offer1.pdf")
    output_path = os.path.join(OUTPUT_FOLDER, "items_offer1.json")
    
    success = extract_items_from_pdf(pdf_path, output_path)
    sys.exit(0 if success else 1)