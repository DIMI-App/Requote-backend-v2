import os
import sys
import json
import openai
import re

try:
    OPENAI_TIMEOUT = int(os.getenv("OPENAI_TIMEOUT", "60"))
except ValueError:
    OPENAI_TIMEOUT = 60

openai.api_key = os.environ.get('OPENAI_API_KEY')

def find_pricing_sections(text):
    """
    Find sections of text that likely contain pricing information
    Returns list of text chunks that contain pricing tables
    """
    # Split by page breaks or large gaps
    # Common page indicators
    page_pattern = r'(?:PAG\.|PAGE|Page|OFFER N\.|Pag\.)\s*\d+'
    
    sections = []
    current_section = []
    
    lines = text.split('\n')
    
    for i, line in enumerate(lines):
        current_section.append(line)
        
        # Check if we hit a page break
        if re.search(page_pattern, line, re.IGNORECASE):
            section_text = '\n'.join(current_section)
            
            # Check if this section has pricing indicators
            has_price = bool(re.search(r'\d+[.,]\d+(?:\s*(?:€|EUR|eur))?', section_text))
            has_table_words = any(word in section_text.lower() for word in 
                                 ['price', 'ціна', 'offer', 'accessories', 'equipment', 'packing'])
            
            if has_price or has_table_words:
                sections.append(section_text)
            
            current_section = []
    
    # Add last section
    if current_section:
        section_text = '\n'.join(current_section)
        if re.search(r'\d+[.,]\d+', section_text):
            sections.append(section_text)
    
    return sections

def extract_items_from_text(text, output_path):
    """Extract structured items from raw offer text.

    Returns
    -------
    tuple[bool, dict | None]
        ``(True, None)`` if extraction succeeds, otherwise ``(False, error)``
        where ``error`` contains diagnostic information for the caller.
    """
    try:
        print(f"📄 Total text length: {len(text)} characters")
        
        # Find pricing sections
        print("\n🔍 Finding pricing sections...")
        pricing_sections = find_pricing_sections(text)
        
        print(f"✅ Found {len(pricing_sections)} sections with pricing")
        
        # If we found sections, use them; otherwise use beginning of document
        if pricing_sections:
            # Combine pricing sections (up to 12000 chars)
            combined_text = '\n\n=== SECTION BREAK ===\n\n'.join(pricing_sections)
            if len(combined_text) > 12000:
                combined_text = combined_text[:12000]
            extraction_text = combined_text
            print(f"   Using {len(extraction_text)} chars from pricing sections")
        else:
            # Fallback: use first 12000 chars
            extraction_text = text[:12000]
            print(f"   No pricing sections found, using first {len(extraction_text)} chars")
        
        prompt = f"""
You are an AI assistant that extracts ALL pricing items from supplier quotations.

CRITICAL INSTRUCTIONS:
- This document contains MULTIPLE pricing tables across different pages
- Extract EVERY item that has a price, no matter which table it's in
- Look for tables with headers like: "Economic Offer", "Accessories", "Packing", "Equipment"
- Even if prices are missing for some items, include them if they're in a pricing table

For each item with pricing information, extract:
- item_name: Full description
- quantity: Quantity (or "" if not specified)
- unit_price: Price per unit (or "" if not specified)
- total_price: Total price (or "" if not specified)
- details: Any additional information

Return ONLY a JSON array:
[
  {{
    "item_name": "...",
    "quantity": "...",
    "unit_price": "...",
    "total_price": "...",
    "details": "..."
  }}
]

If no items found, return empty array: []

Document sections:
{extraction_text}

Extract ALL pricing items. Return ONLY the JSON array.
"""
        
        print("\n🔄 Calling OpenAI to extract items...")
        
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You extract ALL pricing items from quotations. Return only JSON array."},
                {"role": "user", "content": prompt}
            ],
            temperature=0,
            max_tokens=3000,  # Increased for more items
            request_timeout=OPENAI_TIMEOUT
        )
        
        print("📨 Received response from OpenAI")
        
        extracted_json = response.choices[0].message.content.strip()
        
        # Clean markdown wrappers
        if extracted_json.startswith("```json"):
            extracted_json = extracted_json.replace("```json", "").replace("```", "").strip()
        elif extracted_json.startswith("```"):
            extracted_json = extracted_json.replace("```", "").strip()
        
        # Parse JSON
        items = json.loads(extracted_json)
        
        print(f"✅ Validated JSON with {len(items)} items")
        
        if len(items) > 0:
            print(f"   First item: {items[0].get('item_name', 'NO NAME')[:60]}")
        
        # Save in expected format
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        full_data = {
            "items": items,
            "metadata": {
                "sections_analyzed": len(pricing_sections),
                "text_length_analyzed": len(extraction_text)
            }
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(full_data, f, indent=2, ensure_ascii=False)
        
        if os.path.exists(output_path):
            print(f"✅ File created successfully!")
        else:
            print(f"❌ File creation failed")
            return False
        
        print(f"✅ Successfully extracted {len(items)} items")
        return True, None

    except openai.error.Timeout as e:
        print("❌ OpenAI request timed out")
        return False, {
            "type": "timeout",
            "message": str(e) or "OpenAI request timed out",
        }
    except openai.error.OpenAIError as e:
        status = getattr(e, "http_status", None) or getattr(e, "status_code", None)
        error_info = {
            "type": "openai_error",
            "status": status,
            "message": str(e)
        }
        print(f"❌ OpenAI API error: {error_info['message']}")
        if status:
            print(f"   HTTP status: {status}")
        return False, error_info
        return True
        
    except openai.error.Timeout:
        print("❌ OpenAI request timed out")
        return False
    except openai.error.OpenAIError as e:
        print(f"❌ OpenAI API error: {str(e)}")
        return False
    except json.JSONDecodeError as e:
        print(f"❌ JSON parsing error: {str(e)}")
        snippet = extracted_json[:500] if 'extracted_json' in locals() else ''
        if snippet:
            print(f"Response was: {snippet}")
        return False, {
            "type": "json_error",
            "message": str(e),
        }
    except Exception as e:
        print(f"❌ Error during extraction: {str(e)}")
        import traceback
        traceback.print_exc()
        return False, {
            "type": "unexpected",
            "message": str(e),
        }

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
    
    success, error = extract_items_from_text(text, output_json_path)

    if not success and error:
        print(f"❌ Extraction failed: {error.get('message', 'Unknown error')}")

    sys.exit(0 if success else 1)
