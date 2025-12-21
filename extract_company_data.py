import os
import sys
import json
import openai
import base64
from docx import Document
from PIL import Image
import io

openai.api_key = os.environ.get('OPENAI_API_KEY')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def extract_logo_from_docx(docx_path):
    """Extract logo image from DOCX header/body"""
    try:
        doc = Document(docx_path)
        
        # Check for images in document
        images = []
        image_index = 0
        
        # Look in document relationships for images
        for rel in doc.part.rels.values():
            if "image" in rel.target_ref:
                image_index += 1
                image_data = rel.target_part.blob
                
                # Convert to base64
                img_base64 = base64.b64encode(image_data).decode('utf-8')
                
                # Detect format
                if image_data.startswith(b'\xff\xd8\xff'):
                    img_format = 'jpeg'
                elif image_data.startswith(b'\x89PNG'):
                    img_format = 'png'
                else:
                    img_format = 'unknown'
                
                images.append({
                    'index': image_index,
                    'data': img_base64,
                    'format': img_format,
                    'size': len(image_data)
                })
        
        if images:
            # Return first image (usually the logo)
            print(f"✓ Found {len(images)} image(s) in document", flush=True)
            return images[0]
        else:
            print("⚠ No logo image found in document", flush=True)
            return None
            
    except Exception as e:
        print(f"✗ Logo extraction failed: {str(e)}", flush=True)
        return None

def render_docx_page_to_image(docx_path):
    """Render first page of DOCX to image for Vision API analysis"""
    try:
        # Convert DOCX to PDF using LibreOffice
        pdf_path = docx_path.replace('.docx', '_temp.pdf')
        
        result = os.system(f'soffice --headless --convert-to pdf --outdir {os.path.dirname(docx_path)} {docx_path} > /dev/null 2>&1')
        
        if result != 0 or not os.path.exists(pdf_path):
            print("⚠ LibreOffice conversion failed, using alternative method", flush=True)
            return None
        
        # Convert first page of PDF to image
        import fitz  # PyMuPDF
        doc = fitz.open(pdf_path)
        page = doc[0]
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
        img_bytes = pix.tobytes("png")
        doc.close()
        
        # Clean up temp PDF
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        
        img_base64 = base64.b64encode(img_bytes).decode('utf-8')
        print(f"✓ Rendered first page to image ({len(img_base64)} bytes)", flush=True)
        
        return f"data:image/png;base64,{img_base64}"
        
    except Exception as e:
        print(f"⚠ Page rendering failed: {str(e)}", flush=True)
        return None

def extract_company_data_from_offer2(offer2_path, output_path):
    """Extract company branding and information from Offer 2 template"""
    
    try:
        print("=" * 60, flush=True)
        print("EXTRACTING COMPANY DATA FROM OFFER 2", flush=True)
        print("=" * 60, flush=True)
        
        if not openai.api_key:
            print("✗ OPENAI_API_KEY not set", flush=True)
            return False
        
        print(f"Reading template: {offer2_path}", flush=True)
        
        if not os.path.exists(offer2_path):
            print("✗ Template file not found", flush=True)
            return False
        
        # Extract logo image
        logo_data = extract_logo_from_docx(offer2_path)
        
        # Render first page to image for Vision API
        page_image = render_docx_page_to_image(offer2_path)
        
        if not page_image:
            print("⚠ Could not render page, using text extraction only", flush=True)
            # Fallback: extract text from document
            doc = Document(offer2_path)
            text_content = []
            for para in doc.paragraphs[:20]:
                if para.text.strip():
                    text_content.append(para.text.strip())
            
            combined_text = "\n".join(text_content)
            
            # Use text-only GPT analysis
            print("Using text-based extraction...", flush=True)
            
            response = openai.ChatCompletion.create(
                model="gpt-4o",
                messages=[{
                    "role": "user",
                    "content": f"""Extract company information from this quotation template. This is NOT the equipment being sold - this is the COMPANY who is creating the quotation.

Extract ONLY company/seller information:

Template text:
{combined_text[:2000]}

Return ONLY JSON:
{{
  "company_name": "Company legal name",
  "address": "Full address",
  "phone": "Phone number",
  "email": "Email address",
  "website": "Website URL",
  "tax_id": "VAT/Tax ID number",
  "registration_no": "Company registration number",
  "bank_details": {{
    "bank_name": "...",
    "iban": "...",
    "swift": "...",
    "account_holder": "..."
  }},
  "standard_terms": {{
    "delivery": "Standard delivery terms",
    "payment": "Standard payment terms",
    "warranty": "Standard warranty terms"
  }}
}}"""
                }],
                max_tokens=1500,
                temperature=0
            )
            
            extracted_json = response.choices[0].message.content.strip()
            
        else:
            # Use Vision API with rendered page
            print("Using Vision API for extraction...", flush=True)
            
            content = [
                {"type": "text", "text": """Extract COMPANY INFORMATION from this quotation template.

CRITICAL: You are extracting information about the COMPANY/SELLER who creates quotations, NOT the equipment being sold.

Extract:

1. COMPANY IDENTITY:
   - Company legal name
   - Full address (street, city, postal code, country)
   - Phone number(s)
   - Email address(es)
   - Website URL
   - Tax/VAT ID number
   - Company registration number

2. BANK DETAILS:
   - Bank name
   - IBAN
   - SWIFT/BIC code
   - Account holder name

3. STANDARD COMMERCIAL TERMS:
   - Standard delivery terms (e.g., "14 business days")
   - Standard payment terms (e.g., "50% advance, 50% before shipment")
   - Standard warranty terms (e.g., "12 months manufacturer warranty")
   - Any other standard conditions

4. LEGAL/FOOTER INFO:
   - Any legal disclaimers
   - Registration details
   - Certification mentions (ISO, CE, etc.)

IMPORTANT:
- Extract company info from header, footer, and company details sections
- DO NOT extract equipment names, product descriptions, or pricing
- If a field is not present, use empty string ""

Return ONLY JSON:
{
  "company_name": "...",
  "address": "...",
  "phone": "...",
  "email": "...",
  "website": "...",
  "tax_id": "...",
  "registration_no": "...",
  "bank_details": {
    "bank_name": "...",
    "iban": "...",
    "swift": "...",
    "account_holder": "..."
  },
  "standard_terms": {
    "delivery": "...",
    "payment": "...",
    "warranty": "..."
  },
  "legal_info": "..."
}
"""},
                {"type": "image_url", "image_url": {"url": page_image}}
            ]
            
            response = openai.ChatCompletion.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": content}],
                max_tokens=2000,
                temperature=0
            )
            
            extracted_json = response.choices[0].message.content.strip()
        
        # Clean JSON formatting
        if extracted_json.startswith("```json"):
            extracted_json = extracted_json.replace("```json", "").replace("```", "").strip()
        elif extracted_json.startswith("```"):
            extracted_json = extracted_json.replace("```", "").strip()
        
        print("Parsing extracted data...", flush=True)
        company_data = json.loads(extracted_json)
        
        # Add logo data if extracted
        if logo_data:
            company_data['logo'] = {
                'format': logo_data['format'],
                'data': logo_data['data'],
                'size': logo_data['size']
            }
            print(f"✓ Logo included ({logo_data['size']} bytes, {logo_data['format']})", flush=True)
        else:
            company_data['logo'] = None
            print("⚠ No logo found", flush=True)
        
        # Validation
        print("\n" + "=" * 60, flush=True)
        print("EXTRACTED COMPANY DATA", flush=True)
        print("=" * 60, flush=True)
        print(f"Company: {company_data.get('company_name', 'N/A')}", flush=True)
        print(f"Address: {company_data.get('address', 'N/A')[:50]}...", flush=True)
        print(f"Phone: {company_data.get('phone', 'N/A')}", flush=True)
        print(f"Email: {company_data.get('email', 'N/A')}", flush=True)
        print(f"Bank: {company_data.get('bank_details', {}).get('bank_name', 'N/A')}", flush=True)
        print(f"IBAN: {company_data.get('bank_details', {}).get('iban', 'N/A')}", flush=True)
        print("=" * 60, flush=True)
        
        # Save to JSON
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(company_data, f, indent=2, ensure_ascii=False)
        
        print(f"✓ Saved to {output_path}", flush=True)
        print("=" * 60, flush=True)
        
        return True
        
    except Exception as e:
        print(f"✗ FATAL ERROR: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Company Data Extraction Script Started", flush=True)
    
    offer2_path = os.path.join(BASE_DIR, "offer2_template.docx")
    output_path = os.path.join(OUTPUT_FOLDER, "company_data.json")
    
    if not os.path.exists(offer2_path):
        print(f"✗ Template not found at {offer2_path}", flush=True)
        sys.exit(1)
    
    success = extract_company_data_from_offer2(offer2_path, output_path)
    
    if not success:
        print("✗ Extraction failed", flush=True)
        sys.exit(1)
    
    print("✓ COMPLETED SUCCESSFULLY", flush=True)
    sys.exit(0)