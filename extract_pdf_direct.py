import os
import sys
import json
import openai
import fitz  # PyMuPDF
import base64
from io import BytesIO
from PIL import Image

openai.api_key = os.environ.get('OPENAI_API_KEY')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def extract_items_from_pdf(pdf_path, output_path):
    try:
        if not openai.api_key:
            print("ERROR: OPENAI_API_KEY not set")
            return False
        
        print("Reading PDF: " + pdf_path)
        
        # Open PDF with PyMuPDF
        doc = fitz.open(pdf_path)
        print(f"PDF has {len(doc)} pages")
        
        # Convert pages to images
        image_data_list = []
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # Render page to image (72 DPI is fine for text/tables)
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x zoom = 144 DPI
            
            # Convert to PNG bytes
            img_bytes = pix.tobytes("png")
            img_base64 = base64.b64encode(img_bytes).decode('utf-8')
            
            image_data_list.append(f"data:image/png;base64,{img_base64}")
            print(f"  Page {page_num + 1}: converted")
        
        doc.close()
        
        print("Calling OpenAI Vision...")
        
        # Build content with all pages
        content = [
            {"type": "text", "text": """Extract EVERY item with a price or marked "included" from this quotation.

RULES:
1. Extract until "Terms and Conditions" or document end
2. If you see €, $, £, "Included", "Optional" after last item - keep extracting
3. Check ALL sections: Main, Optional, Accessories, Packing, Add-ons
4. Never skip items at end or marked "Included"

Return ONLY JSON:
[{"item_name": "description", "quantity": "1", "unit_price": "€1,000", "total_price": "€1,000", "details": "specs"}]"""}
        ]
        
        # Add all page images
        for img_data in image_data_list:
            content.append({"type": "image_url", "image_url": {"url": img_data}})
        
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": content}],
            max_tokens=4000,
            temperature=0
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
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump({"items": items}, f, indent=2, ensure_ascii=False)
        
        print(f"Saved to {output_path}")
        return True
        
    except Exception as e:
        print(f"ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    pdf_path = os.path.join(UPLOAD_FOLDER, "offer1.pdf")
    output_path = os.path.join(OUTPUT_FOLDER, "items_offer1.json")
    
    if not os.path.exists(pdf_path):
        print("ERROR: PDF not found")
        sys.exit(1)
    
    success = extract_items_from_pdf(pdf_path, output_path)
    
    if not success:
        sys.exit(1)
    
    print("COMPLETED")
    sys.exit(0)