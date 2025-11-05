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
        
        max_chars = 30000
        if len(text) > max_chars:
            print("Text is " + str(len(text)) + " chars, truncating to " + str(max_chars))
            text_to_process = text[:max_chars]
        else:
            text_to_process = text
        
        print(f"Processing {len(text_to_process)} characters")
        
        prompt = """You are an AI specialized in extracting pricing information from business quotations across all industries, languages, and formats.

YOUR TASK: Extract EVERY item that has a price or is marked as "included" from the provided supplier quotation document.

DOCUMENT COMPLETENESS RULES - YOU MUST EXTRACT UNTIL ALL CONDITIONS ARE MET:

1. REACHED END OF PRICING CONTENT
   Stop only when you see: "Terms and Conditions", "Payment Terms", "Delivery Information", "Warranty Information", "Legal Disclaimers", or clear end of document

2. NO MORE PRICE INDICATORS
   If you see ANY of these after your last item, keep extracting: €, $, £, ¥, USD, EUR, GBP, "Included", "Free", "Optional", "price in", numbers in price columns

3. ALL SECTIONS PROCESSED
   Extract from: Main offer, "Optional" sections, "Accessories" sections, "Additional Equipment", "Packing" options, "Add-ons", ANY section with prices

4. NO CONTINUATION INDICATORS
   Keep reading if you see: "See next page", "Continued", "Additional options", "Available as option", page numbers continuing

EXTRACTION PRINCIPLES:

1. SCAN ENTIRE DOCUMENT - Read from beginning to END, check every page and section

2. RECOGNIZE ALL PRICING STRUCTURES:
   - Tables with Description + Price columns
   - Lists with prices (bulleted, numbered)
   - Text blocks with inline pricing
   - Section-based items (Optional, Accessories, Add-ons, Packing)

3. IDENTIFY PRICES IN ANY FORMAT:
   €1,000 | $1,000.00 | £1.000,00 | 5000 (in price columns) | "Included" | "Free" | "TBD"

4. MULTI-LANGUAGE SUPPORT:
   English, Ukrainian, Spanish, German, French, Italian - recognize headers in any language

5. EXTRACT COMPLETE INFO:
   - Full description with model numbers
   - Unit price with currency
   - Quantity (default "1")
   - Total price if shown
   - Notes (optional, included, etc.)

6. NEVER SKIP:
   - Items at document end
   - Items marked "Included" 
   - Items in "optional" sections
   - Items in multiple tables
   - Small-font items

SELF-VALIDATION BEFORE RETURNING:

Q1: Did I read until end of document?
Q2: Any pricing info after my last item? → Go back and extract
Q3: Processed every section with prices?
Q4: Extracted items marked "Included"?
Q5: Checked accessories/add-ons/packing sections?
Q6: Multiple tables? Extracted from all?
Q7: Saw continuation indicators? → Keep extracting

COMPLETION DETECTION - Finished when ALL true:
✓ No more currency/price indicators after last item
✓ No more section headers with pricing
✓ Reached "Terms"/"Conditions"/"Payment"/"Delivery" sections
✓ No tables with price columns remaining

Document text:
""" + text_to_process + """

Return ONLY JSON array:
[{"item_name": "Full description", "quantity": "1", "unit_price": "€1,000.00", "total_price": "€1,000.00", "details": "Model/specs"}]
"""
        
        print("Calling OpenAI...")
        
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo-16k",
            messages=[
                {"role": "system", "content": "Extract ALL pricing data from quotations. Extract until document is complete, not until reaching a count. Validate your work by checking for remaining price indicators. Return only valid JSON."},
                {"role": "user", "content": prompt}
            ],
            temperature=0,
            max_tokens=4000,
            request_timeout=120
        )
        
        print("Received response")
        
        extracted_json = response.choices[0].message.content.strip()
        
        if extracted_json.startswith("```json"):
            extracted_json = extracted_json.replace("```json", "").replace("```", "").strip()
        elif extracted_json.startswith("```"):
            extracted_json = extracted_json.replace("```", "").strip()
        
        items = json.loads(extracted_json)
        
        print(f"Extracted {len(items)} items")
        
        if len(items) == 0:
            print("ERROR: No items extracted")
            return False
        
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        full_data = {"items": items}
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(full_data, f, indent=2, ensure_ascii=False)
        
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print(f"File created: {file_size} bytes")
            if len(items) > 0:
                print(f"First item: {items[0].get('item_name', 'N/A')[:60]}")
                print(f"Last item: {items[-1].get('item_name', 'N/A')[:60]}")
        else:
            print("ERROR: File not created")
            return False
        
        return True
        
    except Exception as e:
        print(f"ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("ITEM EXTRACTION - SV6 COMPLETION-BASED")
    print("=" * 60)
    
    if len(sys.argv) != 3:
        print(f"ERROR: Wrong arguments. Got {len(sys.argv)} arguments: {sys.argv}")
        print("Usage: python extract_items.py <input> <output>")
        sys.exit(1)
    
    input_text_path = sys.argv[1]
    output_json_path = sys.argv[2]
    
    print(f"Input: {input_text_path}")
    print(f"Output: {output_json_path}")
    
    if not os.path.exists(input_text_path):
        print("ERROR: Input file not found")
        sys.exit(1)
    
    with open(input_text_path, 'r', encoding='utf-8') as f:
        text = f.read()
    
    print(f"Read {len(text)} characters")
    
    success = extract_items_from_text(text, output_json_path)
    
    if not success:
        print("Extraction failed")
        sys.exit(1)
    
    print("=" * 60)
    print("COMPLETED")
    print("=" * 60)
    sys.exit(0)