import os
from google.cloud import documentai_v1 as documentai
from google.oauth2 import service_account
import json

# Get absolute paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')

# Ensure output folder exists
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Google Cloud credentials from environment variable
PROJECT_ID = os.environ.get('GCP_PROJECT_ID')
LOCATION = 'us'
PROCESSOR_ID = os.environ.get('GCP_PROCESSOR_ID')
CREDENTIALS_JSON = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS_JSON')

def get_document_ai_client():
    """Initialize Document AI client with credentials from environment"""
    if CREDENTIALS_JSON:
        credentials_dict = json.loads(CREDENTIALS_JSON)
        credentials = service_account.Credentials.from_service_account_info(credentials_dict)
        return documentai.DocumentProcessorServiceClient(credentials=credentials)
    return documentai.DocumentProcessorServiceClient()

def process_pdf_with_documentai(pdf_path, output_text_path):
    """
    Process PDF with Google Document AI and save extracted text
    """
    try:
        print(f"🔍 Processing PDF: {pdf_path}")
        
        # Initialize client
        client = get_document_ai_client()
        
        # Read the file
        with open(pdf_path, 'rb') as f:
            file_content = f.read()
        
        print(f"✅ Read {len(file_content)} bytes from PDF")
        
        # Configure the process request
        name = f"projects/{PROJECT_ID}/locations/{LOCATION}/processors/{PROCESSOR_ID}"
        
        # Create the document
        raw_document = documentai.RawDocument(
            content=file_content,
            mime_type='application/pdf'
        )
        
        # Create the request
        request = documentai.ProcessRequest(
            name=name,
            raw_document=raw_document
        )
        
        print("🔄 Calling Document AI API...")
        
        # Process the document
        result = client.process_document(request=request)
        document = result.document
        
        extracted_text = document.text
        
        print(f"✅ Extracted {len(extracted_text)} characters")
        
        # Save extracted text
        with open(output_text_path, 'w', encoding='utf-8') as f:
            f.write(extracted_text)
        
        print(f"✅ Saved extracted text to: {output_text_path}")
        
        # Verify file was created
        if os.path.exists(output_text_path):
            file_size = os.path.getsize(output_text_path)
            print(f"✅ File verified! Size: {file_size} bytes")
        else:
            print(f"❌ ERROR: File was not created at {output_text_path}")
            return False
        
        return True
        
    except Exception as e:
        print(f"❌ Error processing PDF: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("STARTING DOCUMENT AI PROCESSING")
    print("=" * 60)
    
    # Define paths
    pdf_path = os.path.join(UPLOAD_FOLDER, 'offer1.pdf')
    output_text_path = os.path.join(OUTPUT_FOLDER, 'extracted_text.txt')
    
    print(f"📖 Input PDF: {pdf_path}")
    print(f"💾 Output text: {output_text_path}")
    
    # Check if PDF exists
    if not os.path.exists(pdf_path):
        print(f"❌ PDF not found: {pdf_path}")
        exit(1)
    
    # Process the PDF
    success = process_pdf_with_documentai(pdf_path, output_text_path)
    
    if not success:
        print("❌ Processing failed")
        exit(1)
    
    print("=" * 60)
    print("✅ DOCUMENT AI PROCESSING COMPLETED")
    print("=" * 60)
    exit(0)