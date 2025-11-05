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
        
        # Increased limit for longer documents
        max_chars = 30000
        if len(text) > max_chars:
            print("Text is " + str(len(text)) + " chars, truncating to " + str(max_chars))
            text_to_process = text[:max_chars]
        else:
            text_to_process = text
        
        print(f"Processing {len(text_to_process)} characters")
        
        # COMPLETION-BASED UNIVERSAL PROMPT
        prompt = """You are an AI specialized in extracting pricing information from business quotations across all industries, languages, and formats.

YOUR TASK: Extract EVERY item that has a price or is marked as "included" from the provided supplier quotation document.

═══════════════════════════════════════════════════════════
CRITICAL: DOCUMENT COMPLETENESS RULES
═══════════════════════════════════════════════════════════

YOU MUST EXTRACT UNTIL YOU MEET ALL OF THESE CONDITIONS:

1. ✓ REACHED END OF PRICING CONTENT
   Stop only when you see sections like:
   - "Terms and Conditions" / "Terms of Sale" / "General Terms"
   - "Payment Terms" / "Payment Conditions"
   - "Delivery Information" / "Shipping Terms"
   - "Warranty Information" (non-priced warranty text)
   - "Legal Disclaimers" / "Liability"
   - Clear end of document

2. ✓ NO MORE PRICE INDICATORS
   If you see ANY of these after your last item, keep extracting:
   - Currency symbols: €, $, £, ¥, ₴, USD, EUR, GBP
   - Price formats: "1,000.00" or "1.000,00" or "1 000"
   - Words: "Included", "Free", "No charge", "Optional", "Extra charge", "Additional cost"
   - Numbers in price-like columns or tables
   - Phrases: "price in €", "cost of", "amount"

3. ✓ ALL SECTIONS PROCESSED
   Check you extracted from ALL sections containing prices:
   - Main offer/economic section (always present)
   - "Optional" sections
   - "Accessories" sections
   - "Additional Equipment" sections
   - "Equipments" or "Equipment for" sections
   - "Packing" or "Shipping" or "Transport" options
   - "Add-ons" or "Extras" or "Supplementary"
   - "General Accessories" sections
   - ANY section with a table containing prices

4. ✓ NO CONTINUATION INDICATORS
   If you see these phrases, keep reading and extracting:
   - "See next page" / "Continued on next page"
   - "Additional options below"
   - "Available as option" / "Optional items"
   - Section headers suggesting more content
   - Page numbers continuing (Page 1 of 5, etc.)
   - "More items available"

═══════════════════════════════════════════════════════════
EXTRACTION PRINCIPLES
═══════════════════════════════════════════════════════════

1. SCAN THE ENTIRE DOCUMENT
   - Read from beginning to END
   - Don't stop after finding the first table
   - Check every page, every section
   - Look for items at the very end too

2. RECOGNIZE ALL PRICING STRUCTURES
   
   A) TABLES with columns:
      - Description + Price
      - Item + Cost + Quantity
      - Product + Unit Price + Total
      - No. + Name + Amount
   
   B) LISTS with prices:
      - Bulleted items with prices
      - Numbered lists (1., 2., 3...) with costs
      - Description followed by price on same line
   
   C) TEXT BLOCKS:
      - "Item X: €Y"
      - "Product A costs $B"
      - Inline pricing in paragraphs
   
   D) SECTION-BASED:
      - Items under "Optional Items"
      - Items under "Accessories"
      - Items under "Add-ons"
      - Items under "Packing Options"

3. IDENTIFY PRICES IN ANY FORMAT
   Valid price indicators:
   - €1,000 | €1.000 | €1 000
   - $1,000.00 | 1000 USD
   - £1.000,00 | GBP 1000
   - Just numbers in price columns: 5000
   - Text: "Included" | "Free" | "No charge" | "Complimentary" | "TBD" | "On request"

4. HANDLE MULTI-LANGUAGE DOCUMENTS
   Recognize headers in ANY language:
   - English: "Description", "Price", "Quantity", "Total", "Optional"
   - Ukrainian: "Опис", "Ціна", "Кількість", "Сума"
   - Spanish: "Descripción", "Precio", "Cantidad"
   - German: "Beschreibung", "Preis", "Menge"
   - French: "Description", "Prix", "Quantité"
   - Italian: "Descrizione", "Prezzo", "Quantità"

5. EXTRACT COMPLETE INFORMATION
   For each item capture:
   - Full description (including model numbers, specifications)
   - Unit price (with currency)
   - Quantity (default to "1" if not specified)
   - Total price (if shown)
   - Any notes (optional, included, required, etc.)

6. COMMON SECTIONS TO CHECK
   Always look for items in:
   - Main pricing table
   - Optional items section
   - Accessories tables
   - Equipment add-ons
   - Service packages
   - Warranties (priced ones)
   - Shipping/packing options
   - Installation services
   - Training or documentation
   - Spare parts lists

7. NEVER SKIP
   - Items at document end
   - Items on last pages
   - Items marked "Included" (these count!)
   - Items with price "0"
   - Items in multiple tables
   - Items in appendices
   - Small-font items
   - Items in "optional" sections

═══════════════════════════════════════════════════════════
SELF-VALIDATION CHECKLIST (MANDATORY)
═══════════════════════════════════════════════════════════

Before returning your JSON, verify ALL of these:

Q1: Did I read until the end of the document?
    → Check: Is there text after my last item?
    → If YES: Review that text for prices

Q2: Is there ANY pricing information after my last extracted item?
    → Look for: €, $, £, numbers, "Included", price columns
    → If YES: Go back and extract those items

Q3: Did I process every section header that might contain items?
    → Look for: "Optional", "Accessories", "Additional", "Packing", "Equipment"
    → If found: Extract all items from those sections

Q4: Did I extract items marked as "Included" or "Free"?
    → These are valid items even without numeric prices
    → If missed: Add them to extraction

Q5: Did I check for accessories/add-ons/packing sections?
    → These sections often appear AFTER the main table
    → If not checked: Scan document again

Q6: Are there multiple tables in the document?
    → If YES: Extract from ALL tables, not just the first one

Q7: Did I see continuation indicators?
    → Look for: "See next page", "Continued", "Additional options"
    → If YES: Continue extracting

═══════════════════════════════════════════════════════════
COMPLETION DETECTION
═══════════════════════════════════════════════════════════

You have finished extracting when ALL of these are true:

✓ No more currency symbols or numbers in price format after last item
✓ No more section headers indicating pricing content
✓ Reached sections like "Terms", "Conditions", "Payment", "Delivery", "Warranty"
✓ No tables with price columns remaining
✓ Document ends or only legal/administrative text remains

═══════════════════════════════════════════════════════════

Document text:
""" + text_to_process + """

═══════════════════════════════════════════════════════════

Return ONLY a JSON array with this exact structure:
[{"item_name": "Full item description", "quantity": "1", "unit_price": "€1,000.00", "total_price": "€1,000.00", "details": "Model numbers and specs"}]

IMPORTANT REMINDERS:
- Extract from ENTIRE document, not just first section
- Include items marked "Included" or "Free"
- Check for "optional" and "accessories" sections
- Don't stop until you reach non-pricing content
- Quality over speed - completeness is critical
"""
        
        print("Calling OpenAI with COMPLETION-BASED UNIVERSAL PROMPT...")
        print("Using self-validation approach - no arbitrary item counts")
        
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo-16k",
            messages=[
                {
                    "role": "system", 
                    "content": "You are an expert at extracting ALL pricing data from quotations. You extract until the document is complete, not until you reach a certain count. You validate your own work by checking for remaining price indicators. Return only valid JSON."
                },
                {
                    "role": "user", 
                    "content": prompt
                }
            ],
            temperature=0,
            max_tokens=4000,
            request_timeout=120
        )
        
        print("Received response from OpenAI")
        
        extracted_json = response.choices[0].message.content.strip()
        
        # Clean up JSON formatting
        if extracted_json.startswith("```json"):
            extracted_json = extracted_json.replace("```json", "").replace("```", "").strip()
        elif extracted_json.startswith("```"):
            extracted_json = extracted_json.replace("```", "").strip()
        
        items = json.loads(extracted_json)
        
        print("✓ Validated JSON structure")
        print(f"✓ Extracted {len(items)} items")
        
        # Quality checks (not arbitrary counts, but logical validation)
        if len(items) == 0:
            print("✗ ERROR: No items extracted - this is likely wrong")
            return False
        
        if len(items) == 1:
            print("⚠ WARNING: Only 1 item extracted - verify this is a single-item quotation")
        
        # Check for price diversity (good indicator of completeness)
        prices = []
        included_count = 0
        for item in items:
            price = str(item.get('unit_price', '')).lower()
            if 'included' in price or 'free' in price:
                included_count += 1
            else:
                prices.append(price)
        
        print(f"  - Items with numeric prices: {len(prices)}")
        print(f"  - Items marked 'Included/Free': {included_count}")
        
        # Save to file
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        full_data = {"items": items}
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(full_data, f, indent=2, ensure_ascii=False)
        
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print(f"✓ File created: {file_size} bytes")
            
            if len(items) > 0:
                first_item = items[0].get('item_name', 'N/A')[:60]
                last_item = items[-1].get('item_name', 'N/A')[:60]
                print(f"  - First item: {first_item}")
                print(f"  - Last item: {last_item}")
        else:
            print("✗ ERROR: File not created")
            return False
        
        print(f"✓ Successfully extracted {len(items)} items using COMPLETION-BASED approach")
        return True
        
    except json.JSONDecodeError as e:
        print(f"✗ JSON parsing error: {str(e)}")
        print("OpenAI response was not valid JSON")
        return False
    except Exception as e:
        print(f"✗ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("ITEM EXTRACTION - SV6 COMPLETION-BASED")
    print("=" * 60)
    
    if len(sys.argv) != 3:
        print("✗ ERROR: Wrong arguments")
        print("Usage: python extract_items.py <input> <output>")
        sys.exit(1)
    
    input_text_path = sys.argv[1]
    output_json_path = sys.argv[2]
    
    print(f"Input:  {input_text_path}")
    print(f"Output: {output_json_path}")
    
    if not os.path.exists(input_text_path):
        print("✗ ERROR: Input file not found")
        sys.exit(1)
    
    with open(input_text_path, 'r', encoding='utf-8') as f:
        text = f.read()
    
    print(f"✓ Read {len(text)} characters from input")
    
    success = extract_items_from_text(text, output_json_path)
    
    if not success:
        print("\n✗ Extraction failed")
        sys.exit(1)
    
    print("\n" + "=" * 60)
    print("EXTRACTION COMPLETED SUCCESSFULLY")
    print("=" * 60)
    sys.exit(0)