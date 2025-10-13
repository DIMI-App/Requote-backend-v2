import os
import sys
import json
import openai

openai.api_key = os.environ.get('OPENAI_API_KEY')

def extract_items_from_text(text, output_path):
    try:
        prompt = f"""
You are an AI assistant that extracts structured data from supplier quotations.

Given the following quotation text, extract all items/products with their details.

For each item, extract:
- Item name/description
- Quantity
- Unit price
- Total price (if available)
- Any other relevant details (SKU, part number, etc.)

Return the data as a JSON array with this structure:
[
  {{
    "item_name": "Product name",
    "quantity": "number with unit",
    "unit_price": "price",
    "total_price": "price",
    "details": "any additional info"
  }}
]

Quotation text:
{text[:12000]}

Return ONLY the JSON array, no additional text.
"""
        
        print("🔄 Calling OpenAI to extract items...")
        
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a data extraction assistant. Return only valid JSON."},
                {"role": "user", "content": prompt}
            ],
            temperature=0,
            max_tokens=2000
        )
        
        print("📨 Received response from OpenAI")
        
        extracted_json = response.choices[0].message.content.strip()
        
        if extracted_json.startswith("```json"):
            extracted_json = extracted_json.replace("```json", "").replace("```", "").strip()
        elif extracted_json.startswith("```"):
            extracted_json = extracted_json.replace("```", "").strip()
        
        items = json.loads(extracted_json)
        
        print(f"✅ Validated JSON with {len(items)} items")
        
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        # Save as simple items array for compatibility
        full_data = {
            "items": items
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(full_data, f, indent=2, ensure_ascii=False)
        
        if os.path.exists(output_path):
            print(f"✅ File created successfully!")
        
        print(f"✅ Successfully extracted {len(items)} items")
        return True
        
    except Exception as e:
        print(f"❌ Error during extraction: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python extract_items.py <input_text_path> <output_json_path>")
        sys.exit(1)
    
    input_text_path = sys.argv[1]
    output_json_path = sys.argv[2]
    
    if not os.path.exists(input_text_path):
        print(f"❌ Input file not found: {input_text_path}")
        sys.exit(1)
    
    with open(input_text_path, 'r', encoding='utf-8') as f:
        text = f.read()
    
    print(f"✅ Read {len(text)} characters from input file")
    
    success = extract_items_from_text(text, output_json_path)
    
    sys.exit(0 if success else 1)