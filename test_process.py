import os
import sys
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
    try:
        if not CREDENTIALS_JSON:
            print("❌ ERROR: GOOGLE_APPLICATION_CREDENTIALS_JSON not found in environment")
            sys.exit(1)
        
        print("📋 Parsing Google Cloud credentials...")
        credentials_dict = json.loads(CREDENTIALS_JSON)
        
        print("✅ Credentials parsed successfully")
        print(f"📧 Service account email: {credentials_dict.get('client_email', 'N/A')}")
        
        credentials = service_account.Credentials.from_service_account_info(credentials_dict)
        client = documentai.DocumentProcessorServiceClient(credentials=credentials)
        
        print("✅ Document AI client initialized")
        return client
        
    except json.JSONDecodeError as e:
        print(f"❌ ERROR: Invalid JSON in credentials: {str(e)}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ ERROR: Failed to initialize client: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

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
        print(f"📍 Processor path: {name}")
        
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