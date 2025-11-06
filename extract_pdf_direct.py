import os
import sys
import json
import openai
import fitz
import base64

openai.api_key = os.environ.get('OPENAI_API_KEY')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def extract_items_from_pdf(pdf_path, output_path):
    try:
        print("=== STARTING EXTRACTION ===", flush=True)
        
        if not openai.api_key:
            print("ERROR: OPENAI_API_KEY not set", flush=True)
            return False
        
        print("OpenAI key found", flush=True)
        print("Reading PDF: " + pdf_path, flush=True)
        
        if not os.path.exists(pdf_path):
            print("ERROR: PDF file not found", flush=True)
            return False
        
        print("Opening PDF with PyMuPDF...", flush=True)
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        print(f"PDF has {total_pages} pages", flush=True)
        
        # Process ALL pages (up to 15 for typical quotes)
        max_pages = min(15, total_pages)
        print(f"Processing first {max_pages} pages", flush=True)
        
        image_data_list = []
        for page_num in range(max_pages):
            print(f"Converting page {page_num + 1}...", flush=True)
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            img_bytes = pix.tobytes("png")
            img_base64 = base64.b64encode(img_bytes).decode('utf-8')
            image_data_list.append(f"data:image/png;base64,{img_base64}")
            print(f"Page {page_num + 1}: converted ({len(img_base64)} bytes)", flush=True)
        
        doc.close()
        print("All pages converted", flush=True)
        
        print("Building OpenAI request...", flush=True)
        
        content = [
            {"type": "text", "text": """Extract EVERY item with a price or marked "Included" from this quotation.

CRITICAL RULES:
1. Extract from ALL PAGES - scan entire document
2. Extract items from these sections:
   - Main equipment/machinery
   - Economic Offer table (main pricing table)
   - Format Changes
   - Accessories sections
   - Further Options
   - Packing options
3. For each item, capture:
   - Category/section name (e.g., "FORMAT CHANGES", "ACCESSORIES", "CAN FILLER SANITATION")
   - Item description (full text)
   - Quantity (default "1" if not shown)
   - Unit price - Use ONE of these THREE states:
     * Numeric price: "€324.400,00" or "15.400,00" (keep exact format with dots/commas)
     * Included: "Included" (when text says "Included" or price is 0)
     * To be quoted: "On request" (when text says "Can be offered", "To be quoted", "Please inquire", "On request", or similar)
   - Total price (same format as unit price)
4. NEVER mark "Can be offered" or "To be quoted" items as "Included" - use "On request" instead
5. Preserve thousand separators exactly as shown (€324.400,00 or €15.400,00)
6. Continue extracting until you see "GENERAL SALE TERMS" or end of document

Return ONLY JSON array:
[{
  "category": "Main Equipment",
  "item_name": "CAN ISO 20/2 S - clock wisely running direction",
  "quantity": "1",
  "unit_price": "€324.400,00",
  "total_price": "€324.400,00",
  "details": "Based on one size of 0,33L standard aluminium can including Rolls kit 1st and 2nd operation + chuck in stainless steel"
},
{
  "category": "ACCESSORIES",
  "item_name": "FEEDING PUMP",
  "quantity": "1",
  "unit_price": "On request",
  "total_price": "On request",
  "details": "Can be offered according with the product to be filled"
}]

PRICE STATE EXAMPLES:
- "€324.400,00" → unit_price: "€324.400,00"
- "Included" → unit_price: "Included"
- "Can be offered according..." → unit_price: "On request"
- "To be quoted" → unit_price: "On request"
- "Please inquire" → unit_price: "On request"
"""}
        ]
        
        for img_data in image_data_list:
            content.append({"type": "image_url", "image_url": {"url": img_data}})
        
        print("Calling OpenAI Vision API...", flush=True)
        
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": content}],
            max_tokens=6000,
            temperature=0
        )
        
        print("Received response from OpenAI", flush=True)
        
        extracted_json = response.choices[0].message.content.strip()
        
        if extracted_json.startswith("```json"):
            extracted_json = extracted_json.replace("```json", "").replace("```", "").strip()
        elif extracted_json.startswith("```"):
            extracted_json = extracted_json.replace("```", "").strip()
        
        print("Parsing JSON...", flush=True)
        items = json.loads(extracted_json)
        
        print(f"Extracted {len(items)} items", flush=True)
        
        if len(items) == 0:
            print("ERROR: No items extracted", flush=True)
            return False
        
        # Group items by category
        categories = {}
        for item in items:
            cat = item.get("category", "Main Items")
            if cat not in categories:
                categories[cat] = []
            categories[cat].append(item)
        
        print(f"Found {len(categories)} categories:", flush=True)
        for cat, cat_items in categories.items():
            print(f"  - {cat}: {len(cat_items)} items", flush=True)
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump({"items": items, "categories": list(categories.keys())}, f, indent=2, ensure_ascii=False)
        
        print(f"Saved to {output_path}", flush=True)
        print("=== EXTRACTION COMPLETED ===", flush=True)
        return True
        
    except Exception as e:
        print(f"FATAL ERROR: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Script started", flush=True)
    pdf_path = os.path.join(UPLOAD_FOLDER, "offer1.pdf")
    output_path = os.path.join(OUTPUT_FOLDER, "items_offer1.json")
    
    if not os.path.exists(pdf_path):
        print("ERROR: PDF not found at " + pdf_path, flush=True)
        sys.exit(1)
    
    success = extract_items_from_pdf(pdf_path, output_path)
    
    if not success:
        print("Extraction failed", flush=True)
        sys.exit(1)
    
    print("COMPLETED SUCCESSFULLY", flush=True)
    sys.exit(0)