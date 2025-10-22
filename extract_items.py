import os
import sys
import json
import openai

openai.api_key = os.environ.get('OPENAI_API_KEY')

def extract_items_from_text(text, output_path):
    try:
        # Truncate if too long
        max_chars = 15000
        if len(text) > max_chars:
            print(f"⚠️  Text is {len(text)} chars, truncating to {max_chars}")
            text_to_process = text[:max_chars]
        else:
            text_to_process = text
        
        prompt = f"""
You are an AI that extracts pricing data from supplier quotations for industrial equipment.

TASK: Extract ALL items with prices from the document.

LOOK FOR:
1. Tables with "description" and "price in €" columns
2. Equipment names (like "Automatic rotary rinsing machine TECNA MC24")
3. Optional accessories and their prices
4. Total amounts and line items

EXTRACT for each item:
- item_name: Full equipment/item description
- quantity: Number (use "1" if not specified)
- unit_price: Price with € symbol (e.g., "€270,000")
- total_price: Total if shown, otherwise same as unit_price
- details: Model numbers, specifications, features

RULES:
- Include main equipment AND optional items
- If a row has a price, extract it
- Keep equipment model numbers (like "T24C28S4-VN6")
- Include accessories even if marked "optional"
- Format: Clean JSON array only

Document:
{text_to_process}

Return JSON only:
[
  {{"item_name": "...", "quantity": "1", "unit_price": "€...", "total_price": "€...", "details": "..."}}
]
"""
        
        print("🔄 Calling OpenAI to extract items...")
        
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo-16k",
            messages=[
                {"role": "system", "content": "You extract pricing data from industrial equipment quotations. Return only valid JSON arrays."},
                {"role": "user", "content": prompt}
            ],
            temperature=0,
            max_tokens=4000
        )
        
        print("📨 Received response from OpenAI")
        
        extracted_json = response.choices[0].message.content.strip()
        
        print(f"📝 Raw response length: {len(extracted_json)} characters")
        
        # Clean up markdown code blocks
        if extracted_json.startswith("```json"):
            print("🔧 Removed ```json wrapper")
            extracted_json = extracted_json.replace("```json", "").replace("```", "").strip()
        elif extracted_json.startswith("```"):
            print("🔧 Removed ``` wrapper")
            extracted_json = extracted_json.replace("```", "").strip()
        
        items = json.loads(extracted_json)
        
        print(f"✅ Validated JSON with {len(items)} items")
        
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
            print(f"📁 Ensured directory exists: {output_dir}")
        
        full_data = {"items": items}
        
        print(f"💾 Saving to: {output_path}")
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(full_data, f, indent=2, ensure_ascii=False)
        
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print(f"✅ File created successfully! Size: {file_size} bytes")
            
            # Show preview of first item
            if len(items) > 0:
                print(f"📋 Preview of first item:")
                print(f"   Name: {items[0].get('item_name', 'N/A')[:50]}...")
                print(f"   Price: {items[0].get('unit_price', 'N/A')}")
        else:
            print(f"❌ ERROR: File was not created at {output_path}")
            return False
        
        print(f"✅ Successfully extracted {len(items)} items")
        return True
        
    except json.JSONDecodeError as e:
        print(f"❌ JSON parsing error: {str(e)}")
        print(f"Raw response preview: {extracted_json[:500]}...")
        return False
    except Exception as e:
        print(f"❌ Error during extraction: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("STARTING ITEM EXTRACTION (Day 13 - GPT-3.5-turbo-16k)")
    print("=" * 60)
    
    if len(sys.argv) != 3:
        print("❌ ERROR: Wrong number of arguments")
        print("Usage: python extract_items.py <input_text_path> <output_json_path>")
        sys.exit(1)
    
    input_text_path = sys.argv[1]
    output_json_path = sys.argv[2]
    
    print(f"📖 Input file: {input_text_path}")
    print(f"💾 Output file: {output_json_path}")
    
    if not os.path.exists(input_text_path):
        print(f"❌ Input file not found: {input_text_path}")
        sys.exit(1)
    
    with open(input_text_path, 'r', encoding='utf-8') as f:
        text = f.read()
    
    print(f"✅ Read {len(text)} characters from input file")
    
    success = extract_items_from_text(text, output_json_path)
    
    if not success:
        print("❌ Extraction failed")
        sys.exit(1)
    
    print("=" * 60)
    print("✅ EXTRACTION COMPLETED SUCCESSFULLY")
    print("=" * 60)
    sys.exit(0)