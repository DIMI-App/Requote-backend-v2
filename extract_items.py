import os
import sys
import json
import openai

# OpenAI API key from environment
openai.api_key = os.environ.get('OPENAI_API_KEY')

def extract_items_from_text(text, output_path):
    """
    Extract complete structured data from quotation text using OpenAI GPT-3.5
    """
    try:
        # Truncate text if too long (GPT-3.5 limit is ~4000 tokens ≈ 16000 chars)
        max_chars = 14000
        if len(text) > max_chars:
            print(f"⚠️  Text too long ({len(text)} chars), truncating to {max_chars}")
            # Try to truncate at a reasonable point
            text = text[:max_chars] + "\n\n[Document truncated...]"
        
        prompt = f"""
You are an AI assistant that extracts ALL pricing items from supplier quotations.

CRITICAL: This document contains MULTIPLE pricing tables. You MUST extract EVERY item that has a price.

Look for tables with headers like:
- Economic Offer / Main offer
- General Accessories (optional)
- Accessories of the rinsing turret
- Accessories of the filling turret
- Equipments for corker
- Packing options

For EACH item with a price, extract:
{{
  "item_name": "Full description",
  "quantity": "Quantity (or empty string if not specified)",
  "unit_price": "Price per unit",
  "total_price": "Total price (or empty string if not specified)",
  "details": "Any additional info"
}}

Also extract if available:
- technical_specs: Technical specifications
- company_info: Offer number, date, company name

Return JSON:
{{
  "items": [ array of all items ],
  "technical_specs": {{ "title": "...", "content": "..." }},
  "company_info": {{ "offer_number": "...", "date": "...", "company_name": "..." }}
}}

Document:
{text}

Return ONLY valid JSON.
"""
        
        print("🔄 Calling OpenAI API...")
        
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You extract ALL pricing items from quotations. Return only JSON."},
                {"role": "user", "content": prompt}
            ],
            temperature=0,
            max_tokens=3000
        )
        
        print("📨 Received response from OpenAI")
        
        extracted_json = response.choices[0].message.content.strip()
        
        print(f"📝 Response length: {len(extracted_json)} chars")
        
        # Clean markdown wrappers
        if extracted_json.startswith("```json"):
            extracted_json = extracted_json.replace("```json", "").replace("```", "").strip()
        elif extracted_json.startswith("```"):
            extracted_json = extracted_json.replace("```", "").strip()
        
        # Parse JSON
        full_data = json.loads(extracted_json)
        
        # Validate structure
        if "items" not in full_data:
            full_data["items"] = []
        
        items = full_data.get("items", [])
        
        print(f"✅ Extracted data:")
        print(f"   • Items: {len(items)}")
        print(f"   • Technical specs: {'Yes' if full_data.get('technical_specs') else 'No'}")
        print(f"   • Company info: {'Yes' if full_data.get('company_info') else 'No'}")
        
        # Save to file
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(full_data, f, indent=2, ensure_ascii=False)
        
        if os.path.exists(output_path):
            print(f"✅ Saved to: {output_path}")
        else:
            print(f"❌ Failed to create file")
            return False
        
        print(f"✅ Extraction complete: {len(items)} items")
        return True
        
    except json.JSONDecodeError as e:
        print(f"❌ JSON error: {str(e)}")
        return False
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python extract_items.py <input> <output>")
        sys.exit(1)
    
    input_path = sys.argv[1]
    output_path = sys.argv[2]
    
    if not os.path.exists(input_path):
        print(f"❌ Input not found: {input_path}")
        sys.exit(1)
    
    with open(input_path, 'r', encoding='utf-8') as f:
        text = f.read()
    
    print(f"✅ Read {len(text)} chars")
    
    success = extract_items_from_text(text, output_path)
    sys.exit(0 if success else 1)