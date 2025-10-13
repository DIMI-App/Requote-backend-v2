import os
import sys
import json
import openai

# OpenAI API key from environment
openai.api_key = os.environ.get('OPENAI_API_KEY')

def extract_items_from_text(text, output_path):
    """
    Extract complete structured data from quotation text using OpenAI GPT-3.5
    Extracts: items table + technical specs + company info
    """
    try:
        # === ONLY CHANGE: Better prompt that looks for ALL tables ===
        prompt = f"""
You are an AI assistant that extracts ALL pricing items from supplier quotations.

IMPORTANT: This document may contain MULTIPLE pricing tables across many pages.
You MUST extract EVERY item that has a price, from ALL tables in the document.

Common table types:
- Main offer/economic offer
- General accessories
- Optional accessories
- Equipment options
- Packing options
- Additional formats

For EACH item with a price, extract:
- item_name: Full product description
- quantity: Quantity with unit (or "" if not specified)
- unit_price: Unit price
- total_price: Total price (or "" if not specified)
- details: Any additional information

Also extract if available:
- technical_specs: Any technical specifications sections
- company_info: Offer number, date, company name

Return as JSON:
{{
  "items": [
    {{
      "item_name": "...",
      "quantity": "...",
      "unit_price": "...",
      "total_price": "...",
      "details": "..."
    }}
  ],
  "technical_specs": {{
    "title": "...",
    "content": "..."
  }},
  "company_info": {{
    "offer_number": "...",
    "date": "...",
    "company_name": "..."
  }}
}}

Extract ALL items with prices from ALL tables. Do not skip any tables.

Document text:
{text}

Return ONLY the JSON object.
"""
        
        print("🔄 Calling OpenAI to extract complete document data...")
        
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo-16k",  # Use 16k model for long documents
            messages=[
                {"role": "system", "content": "You are a data extraction assistant. Extract ALL pricing items from ALL tables. Return only valid JSON."},
                {"role": "user", "content": prompt}
            ],
            temperature=0,
            max_tokens=4000  # Increased for more items
        )
        
        print("📨 Received response from OpenAI")
        
        # Extract the JSON from response
        extracted_json = response.choices[0].message.content.strip()
        
        print(f"📝 Raw response length: {len(extracted_json)} characters")
        
        # Clean up the response (remove markdown code blocks if present)
        if extracted_json.startswith("```json"):
            print("🔧 Removed ```json wrapper")
            extracted_json = extracted_json.replace("```json", "").replace("```", "").strip()
        elif extracted_json.startswith("```"):
            print("🔧 Removed ``` wrapper")
            extracted_json = extracted_json.replace("```", "").strip()
        
        # Parse to validate JSON
        full_data = json.loads(extracted_json)
        
        # Verify items array exists
        if "items" not in full_data:
            print("⚠️  WARNING: No 'items' key found, creating empty array")
            full_data["items"] = []
        
        items = full_data.get("items", [])
        
        print(f"✅ Validated JSON:")
        print(f"   • Items: {len(items)}")
        print(f"   • Technical specs: {'Yes' if full_data.get('technical_specs') else 'No'}")
        print(f"   • Company info: {'Yes' if full_data.get('company_info') else 'No'}")
        
        # Ensure output directory exists
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
            print(f"📁 Ensured directory exists: {output_dir}")
        
        # Save to file with absolute path
        print(f"💾 Saving to: {output_path}")
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(full_data, f, indent=2, ensure_ascii=False)
        
        # Verify file was created
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print(f"✅ File created successfully! Size: {file_size} bytes")
        else:
            print(f"❌ ERROR: File was not created at {output_path}")
            return False
        
        print(f"✅ Successfully extracted complete document data with {len(items)} items")
        return True
        
    except json.JSONDecodeError as e:
        print(f"❌ JSON parsing error: {str(e)}")
        print(f"Raw response: {extracted_json[:500]}...")
        return False
    except Exception as e:
        print(f"❌ Error during extraction: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("STARTING ENHANCED ITEM EXTRACTION")
    print("=" * 60)
    
    # Accept input and output paths as command-line arguments
    if len(sys.argv) != 3:
        print("❌ ERROR: Wrong number of arguments")
        print("Usage: python extract_items.py <input_text_path> <output_json_path>")
        sys.exit(1)
    
    input_text_path = sys.argv[1]
    output_json_path = sys.argv[2]
    
    print(f"📖 Input file: {input_text_path}")
    print(f"💾 Output file: {output_json_path}")
    
    # Read input text
    if not os.path.exists(input_text_path):
        print(f"❌ Input file not found: {input_text_path}")
        sys.exit(1)
    
    with open(input_text_path, 'r', encoding='utf-8') as f:
        text = f.read()
    
    print(f"✅ Read {len(text)} characters from input file")
    
    # Extract complete document data
    success = extract_items_from_text(text, output_json_path)
    
    if not success:
        print("❌ Extraction failed")
        sys.exit(1)
    
    print("=" * 60)
    print("✅ EXTRACTION COMPLETED SUCCESSFULLY")
    print("=" * 60)
    sys.exit(0)