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
        
        prompt = """Extract all items with prices from this industrial equipment quotation.

For each item, provide:
- item_name: Equipment description
- quantity: Number (use "1" if not specified)
- unit_price: Price with currency symbol
- total_price: Total if shown
- details: Model numbers and specs

Document:
""" + text_to_process + """

Return only JSON array:
[{"item_name": "...", "quantity": "1", "unit_price": "...", "total_price": "...", "details": "..."}]
"""
        
        print("Calling OpenAI...")
        
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "Extract pricing data. Return only JSON."},
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
        
        print("Successfully extracted " + str(len(items)) + " items")
        return True
        
    except Exception as e:
        print("ERROR: " + str(e))
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("ITEM EXTRACTION - Day 13")
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