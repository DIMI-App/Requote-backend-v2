import os
import sys
import json
import tempfile
from google.cloud import documentai_v1 as documentai
from google.api_core.client_options import ClientOptions

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

def setup_credentials():
    """Setup Google Cloud credentials and return temp file path"""
    print("=" * 60)
    print("STEP 1: Setting up Google Cloud credentials")
    print("=" * 60)
    
    if not CREDENTIALS_JSON:
        print("ERROR: GOOGLE_APPLICATION_CREDENTIALS_JSON not found in environment")
        return None
    
    try:
        print("Parsing credentials JSON...")
        credentials_dict = json.loads(CREDENTIALS_JSON)
        print("SUCCESS: Credentials parsed")
        
        service_email = credentials_dict.get('client_email', 'N/A')
        print("Service account: " + service_email)
        
        # Create temporary credentials file
        print("Creating temporary credentials file...")
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.json') as temp_creds:
            json.dump(credentials_dict, temp_creds)
            temp_creds_path = temp_creds.name
        
        print("Temp file created: " + temp_creds_path)
        
        # Set environment variable
        os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = temp_creds_path
        print("Environment variable set")
        
        return temp_creds_path
        
    except json.JSONDecodeError as e:
        print("ERROR: Invalid JSON in credentials: " + str(e))
        return None
    except Exception as e:
        print("ERROR: Failed to setup credentials: " + str(e))
        import traceback
        traceback.print_exc()
        return None

def initialize_client():
    """Initialize Document AI client"""
    print("")
    print("=" * 60)
    print("STEP 2: Initializing Document AI client")
    print("=" * 60)
    
    try:
        # Set regional endpoint
        client_options = ClientOptions(
            api_endpoint=LOCATION + "-documentai.googleapis.com"
        )
        print("API endpoint: " + LOCATION + "-documentai.googleapis.com")
        
        # Create client
        client = documentai.DocumentProcessorServiceClient(client_options=client_options)
        print("SUCCESS: Client initialized")
        
        return client
        
    except Exception as e:
        print("ERROR: Failed to initialize client: " + str(e))
        import traceback
        traceback.print_exc()
        return None

def read_pdf(pdf_path):
    """Read PDF file and return content"""
    print("")
    print("=" * 60)
    print("STEP 3: Reading PDF file")
    print("=" * 60)
    
    print("PDF path: " + pdf_path)
    
    if not os.path.exists(pdf_path):
        print("ERROR: PDF file not found!")
        return None
    
    try:
        with open(pdf_path, 'rb') as f:
            content = f.read()
        
        print("SUCCESS: Read " + str(len(content)) + " bytes")
        return content
        
    except Exception as e:
        print("ERROR: Failed to read PDF: " + str(e))
        return None

def process_with_documentai(client, pdf_content):
    """Process PDF with Document AI"""
    print("")
    print("=" * 60)
    print("STEP 4: Processing with Document AI")
    print("=" * 60)
    
    try:
        # Build processor path
        name = "projects/" + PROJECT_ID + "/locations/" + LOCATION + "/processors/" + PROCESSOR_ID
        print("Processor: " + name)
        
        # Create document
        raw_document = documentai.RawDocument(
            content=pdf_content,
            mime_type='application/pdf'
        )
        
        # Create request
        request = documentai.ProcessRequest(
            name=name,
            raw_document=raw_document
        )
        
        print("Sending request to Document AI...")
        
        # Process document
        result = client.process_document(request=request)
        document = result.document
        
        print("SUCCESS: Document processed")
        print("Extracted text length: " + str(len(document.text)) + " characters")
        
        return document.text
        
    except Exception as e:
        print("ERROR: Document AI processing failed: " + str(e))
        import traceback
        traceback.print_exc()
        return None

def save_extracted_text(text, output_path):
    """Save extracted text to file"""
    print("")
    print("=" * 60)
    print("STEP 5: Saving extracted text")
    print("=" * 60)
    
    print("Output path: " + output_path)
    
    try:
        # Ensure directory exists
        output_dir = os.path.dirname(output_path)
        print("Output directory: " + output_dir)
        print("Directory exists: " + str(os.path.exists(output_dir)))
        
        if not os.path.exists(output_dir):
            print("Creating directory...")
            os.makedirs(output_dir, exist_ok=True)
        
        # Write file
        print("Writing file...")
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text)
        
        print("File write completed")