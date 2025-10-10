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
            print("ERROR: GOOGLE_APPLICATION_CREDENTIALS_JSON not found")
            sys.exit(1)
        
        print("Parsing Google Cloud credentials...")
        credentials_dict = json.loads(CREDENTIALS_JSON)
        
        print("Credentials parsed successfully")
        service_email = credentials_dict.get('client_email', 'N/A')
        print("Service account email: " + service_email)
        
        credentials = service_account.Credentials.from_service_account_info(credentials_dict)
        client = documentai.DocumentProcessorServiceClient(credentials=credentials)
        
        print("Document AI client initialized")
        return client
        
    except json.JSONDecodeError as e:
        print("ERROR: Invalid JSON in credentials: " + str(e))
        sys.exit(1)
    except Exception as e:
        print("ERROR: Failed to initialize client: " + str(e))
        import traceback
        traceback.print_exc()
        sys.exit(1)

def process_pdf_with_documentai(pdf_path, output_text_path):
    """Process PDF with Google Document AI and save extracted text"""
    try:
        print("Processing PDF: " + pdf_path)
        
        # Initialize client
        client = get_document_ai_client()
        
        # Read the file
        with open(pdf_path, 'rb') as f:
            file_content = f.read()
        
        print("Read " + str(len(file_content)) + " bytes from PDF")
        
        # Configure the process request
        name = "projects/" + PROJECT_ID + "/locations/" + LOCATION + "/processors/" + PROCESSOR_ID
        print("Processor path: " + name)
        
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
        
        print("Calling Document AI API...")
        
        # Process the document
        result = client.process_document(request=request)
        document = result.document
        
        extracted_text = document.text
        
        print("Extracted " + str(len(extracted_text)) + " characters")
        
        # Save extracted text
        with open(output_text_path, 'w', encoding='utf-8') as f:
            f.write(extracted_text)
        
        print("Saved extracted text to: " + output_text_path)
        
        # Verify file was created
        if os.path.exists(output_text_path):
            file_size = os.path.getsize(output_text_path)
            print("File verified! Size: " + str(file_size) + " bytes")
        else:
            print("ERROR: File was not created at " + output_text_path)
            return False
        
        return True
        
    except Exception as e:
        print("Error processing PDF: " + str(e))
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("STARTING DOCUMENT AI PROCESSING")
    print("=" * 60)
    
    # Check required environment variables
    if not PROJECT_ID:
        print("ERROR: GCP_PROJECT_ID not set")
        sys.exit(1)
    
    if not PROCESSOR_ID:
        print("ERROR: GCP_PROCESSOR_ID not set")
        sys.exit(1)
    
    print("Project ID: " + PROJECT_ID)
    print("Processor ID: " + PROCESSOR_ID)
    
    # Define paths
    pdf_path = os.path.join(UPLOAD_FOLDER, 'offer1.pdf')
    output_text_path = os.path.join(OUTPUT_FOLDER, 'extracted_text.txt')
    
    print("Input PDF: " + pdf_path)
    print("Output text: " + output_text_path)
    
    # Check if PDF exists
    if not os.path.exists(pdf_path):
        print("PDF not found: " + pdf_path)
        sys.exit(1)
    
    # Process the PDF
    success = process_pdf_with_documentai(pdf_path, output_text_path)
    
    if not success:
        print("Processing failed")
        sys.exit(1)
    
    print("=" * 60)
    print("DOCUMENT AI PROCESSING COMPLETED")
    print("=" * 60)
    sys.exit(0)