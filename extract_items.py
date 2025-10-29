import os
import sys
import json
import openai

openai.api_key = os.environ.get('OPENAI_API_KEY')

def extract_items_from_text(text, output_path):
    try:
        if not openai.api_key:
            print("ERROR: OPENAI_API_KEY not set")
            return False
        
        print("API key found (length: " + str(len(openai.api_key)) + ")")
        
        max_chars = 12000
        if len(text) > max_chars:
            print("Text is " + str(len(text)) + " chars, truncating to " + str(max_chars))
            text_to_process = text[:max_chars]
        else:
            text_to_process = text
        
        # UNIVERSAL PROMPT - Works with ANY quotation format
        prompt = """You are an AI specialized in extracting pricing information from business quotations across all industries, languages, and formats.

YOUR TASK: Extract EVERY item that has a price or is marked as "included" from the provided supplier quotation document.

UNIVERSAL EXTRACTION PRINCIPLES:

1. SCAN THE ENTIRE DOCUMENT
   - Read every page from beginning to end
   - Don't stop after finding the first table
   - Check every section, even those at the very end

2. RECOGNIZE PRICING STRUCTURES
   Look for ANY of these patterns:
   
   A) TABLES with columns like:
      - Item/Description + Price
      - No./Position + Name + Cost
      - Product + Quantity + Unit Price
      - Code + Description + Amount
   
   B) LISTS with prices:
      - Bulleted items with €/$/£ amounts
      - Numbered items (1., 2., 3...) followed by prices
      - Dash-separated items with costs
   
   C) TEXT BLOCKS with pricing:
      - "Item X costs €Y"
      - "Product A: €B"
      - "Service 1 - $X per unit"
   
   D) SECTION HEADERS indicating pricing:
      - "Optional Items"
      - "Accessories" 
      - "Add-ons"
      - "Additional Equipment"
      - "Extras"
      - "Supplementary Items"

3. IDENTIFY PRICES IN ANY FORMAT
   Recognize these as valid prices:
   - With currency symbols: €1,000 | $1,000.00 | £1.000,00 | ¥1000
   - Numbers only (in "price" columns): 1000 | 1.000 | 1,000.00
   - With text: "1000 euros" | "USD 1000"
   - Special indicators: "Included" | "Free" | "No charge" | "Complimentary" | "On request" | "TBD"

4. HANDLE MULTI-LANGUAGE DOCUMENTS
   Column headers can be in ANY language:
   - English: "Description", "Price", "Quantity", "Total"
   - Spanish: "Descripción", "Precio", "Cantidad"
   - German: "Beschreibung", "Preis", "Menge"
   - French: "Description", "Prix", "Quantité"
   - Ukrainian: "Опис", "Ціна", "Кількість"
   - Italian: "Descrizione", "Prezzo", "Quantità"
   
   Identify columns by their POSITION and CONTENT, not just their names.

5. EXTRACT COMPLETE INFORMATION
   For each item, capture:
   - Position/Number (if present)
   - Full description (including model numbers, specs, technical details)
   - Unit price
   - Quantity (if specified, otherwise assume 1)
   - Total price (if calculated)
   - Any special notes (included, optional, required, etc.)

6. COMMON QUOTATION SECTIONS
   Items can appear in these sections:
   - Main equipment/products table (usually near the beginning)
   - Optional items section (middle or end)
   - Accessories list
   - Service packages
   - Warranties or support plans
   - Shipping/packaging options
   - Installation services
   - Training or documentation
   - Spare parts

7. DO NOT SKIP
   - Items at the very end of the document
   - Items in separate tables on different pages
   - Items with price "0" or "Included" (these are important!)
   - Items in footnotes or appendices
   - Items in smaller font or different formatting

Document text:
""" + text_to_process + """

Return ONLY a JSON array with this exact structure:
[{"item_name": "Full item description", "quantity": "1", "unit_price": "€1,000.00", "total_price": "€1,000.00", "details": "Model numbers and specs"}]

VALIDATION CHECKLIST before responding:
- ✓ Did I read the ENTIRE document?
- ✓ Did I check ALL pages?
- ✓ Did I look for items beyond the first table?
- ✓ Did I include items marked "Included"?
- ✓ Did I check sections labeled "optional" or "accessories"?
- ✓ Is my item count reasonable for this document size?
"""
        
        print("Calling OpenAI with UNIVERSAL PROMPT...")
        
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are an expert at extracting ALL pricing data from quotations. Return only valid JSON."},
                {"role": "user", "content": prompt}
            ],
            temperature=0,
            max_tokens=2000,
            request_timeout=60
        )
        
        print("Received response")
        
        extracted_json = response.choices[0].message.content.strip()
        
        if extracted_json.startswith("```json"):
            extracted_json = extracted_json.replace("```json", "").replace("```", "").strip()
        elif extracted_json.startswith("```"):
            extracted_json = extracted_json.replace("```", "").strip()
        
        items = json.loads(extracted_json)
        
        print("Validated " + str(len(items)) + " items")
        
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        full_data = {"items": items}
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(full_data, f, indent=2, ensure_ascii=False)
        
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print("File created: " + str(file_size) + " bytes")
            
            if len(items) > 0:
                first_item_name = items[0].get('item_name', 'N/A')[:60]
                print("First item: " + first_item_name)
        else:
            print("ERROR: File not created")
            return False
        
        print("Successfully extracted " + str(len(items)) + " items using UNIVERSAL PROMPT")
        return True
        
    except Exception as e:
        print("ERROR: " + str(e))
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("ITEM EXTRACTION - SV6 with UNIVERSAL PROMPTS")
    print("=" * 60)
    
    if len(sys.argv) != 3:
        print("ERROR: Wrong arguments")
        print("Usage: python extract_items.py <input> <output>")
        sys.exit(1)
    
    input_text_path = sys.argv[1]
    output_json_path = sys.argv[2]
    
    print("Input: " + input_text_path)
    print("Output: " + output_json_path)
    
    if not os.path.exists(input_text_path):
        print("ERROR: Input file not found")
        sys.exit(1)
    
    with open(input_text_path, 'r', encoding='utf-8') as f:
        text = f.read()
    
    print("Read " + str(len(text)) + " characters")
    
    success = extract_items_from_text(text, output_json_path)
    
    if not success:
        print("Extraction failed")
        sys.exit(1)
    
    print("=" * 60)
    print("COMPLETED")
    print("=" * 60)
    sys.exit(0)