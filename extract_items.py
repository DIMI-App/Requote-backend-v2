import os
import sys
import json
import openai
import re

openai.api_key = os.environ.get('OPENAI_API_KEY')

def split_text_into_pages(text):
    """
    Split extracted text into page chunks
    Assumes pages are separated by page numbers or clear breaks
    """
    # Try to split by page indicators
    page_pattern = r'(?:PAG\.|PAGE|Page)\s*\d+|OFFER N\.\s*\d+'
    
    pages = []
    current_page = []
    
    for line in text.split('\n'):
        # Check if this line indicates a new page
        if re.search(page_pattern, line) and current_page:
            pages.append('\n'.join(current_page))
            current_page = [line]
        else:
            current_page.append(line)
    
    # Add the last page
    if current_page:
        pages.append('\n'.join(current_page))
    
    return pages

def identify_pricing_pages(pages):
    """
    Identify which pages likely contain pricing information
    """
    pricing_keywords = [
        'price', 'eur', '€', 'amount', 'total',
        'offer', 'economic', 'accessories',
        'optional', 'packing'
    ]
    
    pricing_pages = []
    
    for idx, page in enumerate(pages):
        page_lower = page.lower()
        
        # Count pricing keywords
        keyword_count = sum(1 for keyword in pricing_keywords if keyword in page_lower)
        
        # Check for price patterns (numbers with currency)
        price_patterns = len(re.findall(r'\d+[.,]\d+\s*(?:€|EUR|eur)', page))
        
        # If page has pricing indicators, mark it
        if keyword_count >= 2 or price_patterns >= 2:
            pricing_pages.append({
                'index': idx,
                'content': page,
                'keyword_count': keyword_count,
                'price_patterns': price_patterns
            })
            print(f"   Page {idx}: Pricing detected (keywords: {keyword_count}, prices: {price_patterns})")
    
    return pricing_pages

def extract_items_from_page(page_content, page_number):
    """
    Extract items from a single page using OpenAI
    """
    prompt = f"""
You are extracting pricing items from a supplier quotation document.

This is page {page_number} of the document. Extract ALL items that have prices from this page.

IMPORTANT:
- Extract EVERY item with a price, no matter how small
- Look for multiple tables on the same page
- Include optional accessories, equipment, and add-ons
- Capture both main items and sub-items

For each item, extract:
- item_name: Full description
- quantity: Quantity if specified, otherwise ""
- unit_price: Price per unit
- total_price: Total if specified
- category: Type of item (e.g., "Main Equipment", "Accessories", "Packing")

Return ONLY a JSON array:
[
  {{
    "item_name": "...",
    "quantity": "...",
    "unit_price": "...",
    "total_price": "...",
    "category": "..."
  }}
]

If no pricing items found, return an empty array: []

Page content:
{page_content}

Return ONLY the JSON array.
"""
    
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo-16k",  # Use 16k model for longer context
            messages=[
                {"role": "system", "content": "You are a data extraction assistant. Return only valid JSON."},
                {"role": "user", "content": prompt}
            ],
            temperature=0,
            max_tokens=2000
        )
        
        extracted_json = response.choices[0].message.content.strip()
        
        # Clean up response
        if extracted_json.startswith("```json"):
            extracted_json = extracted_json.replace("```json", "").replace("```", "").strip()
        elif extracted_json.startswith("```"):
            extracted_json = extracted_json.replace("```", "").strip()
        
        # Parse JSON
        items = json.loads(extracted_json)
        
        return items if isinstance(items, list) else []
        
    except Exception as e:
        print(f"   ⚠️  Error extracting from page {page_number}: {e}")
        return []

def extract_items_from_text(text, output_path):
    """
    Multi-stage extraction for complex documents with multiple pricing tables
    """
    try:
        print("\n" + "=" * 60)
        print("🔄 MULTI-STAGE EXTRACTION PROCESS")
        print("=" * 60)
        
        # Step 1: Split into pages
        print("\n📄 Step 1: Splitting document into pages...")
        pages = split_text_into_pages(text)
        print(f"   Found {len(pages)} pages")
        
        # Step 2: Identify pricing pages
        print("\n💰 Step 2: Identifying pricing pages...")
        pricing_pages = identify_pricing_pages(pages)
        print(f"   Found {len(pricing_pages)} pages with pricing information")
        
        if len(pricing_pages) == 0:
            print("   ⚠️  No pricing pages detected!")
            print("   Falling back to full document extraction...")
            # Fallback: try to extract from entire document
            pricing_pages = [{'index': 0, 'content': text, 'keyword_count': 0, 'price_patterns': 0}]
        
        # Step 3: Extract from each pricing page
        print("\n🤖 Step 3: Extracting items from each pricing page...")
        all_items = []
        
        for page_info in pricing_pages:
            page_idx = page_info['index']
            page_content = page_info['content']
            
            print(f"\n   Processing page {page_idx}...")
            print(f"   Content length: {len(page_content)} chars")
            
            items = extract_items_from_page(page_content, page_idx)
            
            if items:
                print(f"   ✅ Extracted {len(items)} items from page {page_idx}")
                all_items.extend(items)
            else:
                print(f"   • No items found on page {page_idx}")
        
        # Step 4: Deduplicate items
        print(f"\n🔍 Step 4: Deduplicating items...")
        unique_items = []
        seen_names = set()
        
        for item in all_items:
            item_name = item.get('item_name', '').strip().lower()
            if item_name and item_name not in seen_names:
                unique_items.append(item)
                seen_names.add(item_name)
            else:
                print(f"   • Skipped duplicate: {item_name[:50]}...")
        
        print(f"   Removed {len(all_items) - len(unique_items)} duplicates")
        print(f"   Final item count: {len(unique_items)}")
        
        # Step 5: Create structured output
        print("\n📦 Step 5: Creating structured output...")
        
        full_data = {
            "items": unique_items,
            "metadata": {
                "total_pages": len(pages),
                "pricing_pages": len(pricing_pages),
                "extraction_method": "multi-stage",
                "total_items_found": len(unique_items)
            }
        }
        
        # Save to file
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(full_data, f, indent=2, ensure_ascii=False)
        
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print(f"\n✅ SUCCESS! Saved to: {output_path}")
            print(f"   File size: {file_size} bytes")
        else:
            print(f"\n❌ ERROR: File not created")
            return False
        
        print("\n" + "=" * 60)
        print("📊 EXTRACTION SUMMARY")
        print("=" * 60)
        print(f"   • Total items extracted: {len(unique_items)}")
        print(f"   • Pages analyzed: {len(pricing_pages)}")
        print(f"   • Categories found: {len(set(item.get('category', 'Unknown') for item in unique_items))}")
        print("=" * 60)
        
        return True
        
    except Exception as e:
        print(f"\n❌ EXTRACTION FAILED: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("🚀 MULTI-STAGE ITEM EXTRACTION")
    print("=" * 60)
    
    if len(sys.argv) != 3:
        print("❌ ERROR: Wrong number of arguments")
        print("Usage: python extract_items.py <input_text_path> <output_json_path>")
        sys.exit(1)
    
    input_text_path = sys.argv[1]
    output_json_path = sys.argv[2]
    
    print(f"\n📖 Input: {input_text_path}")
    print(f"💾 Output: {output_json_path}")
    
    if not os.path.exists(input_text_path):
        print(f"\n❌ Input file not found: {input_text_path}")
        sys.exit(1)
    
    with open(input_text_path, 'r', encoding='utf-8') as f:
        text = f.read()
    
    print(f"\n✅ Read {len(text)} characters from input")
    
    success = extract_items_from_text(text, output_json_path)
    
    sys.exit(0 if success else 1)