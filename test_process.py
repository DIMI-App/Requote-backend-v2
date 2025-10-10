import os
import sys
import json
from google.cloud import documentai_v1 as documentai
from google.oauth2 import service_account
from google.api_core.client_options import ClientOptions

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

PROJECT_ID = os.environ.get('GCP_PROJECT_ID')
LOCATION = 'us'
PROCESSOR_ID = os.environ.get('GCP_PROCESSOR_ID')
CREDENTIALS_JSON = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS_JSON')

def main():
    print("Starting Document AI processing...")
    
    if not PROJECT_ID or not PROCESSOR_ID or not CREDENTIALS_JSON:
        print("ERROR: Missing environment variables")
        return False
    
    print("Project ID: " + PROJECT_ID)
    print("Processor ID: " + PROCESSOR_ID)
    
    try:
        # Parse credentials
        print("Parsing credentials...")
        creds_dict = json.loads(CREDENTIALS_JSON)
        print("Service account: " + creds_dict.get('client_email', 'N/A'))
        
        # Create credentials object directly (no temp file)
        print("Creating credentials object...")
        credentials = service_account.Credentials.from_service_account_info(creds_dict)
        print("Credentials created")
        
        # Initialize client with credentials
        print("Initializing Document AI client...")
        client_options = ClientOptions(api_endpoint=LOCATION + "-documentai.googleapis.com")
        client = documentai.DocumentProcessorServiceClient(
            credentials=credentials,
            client_options=client_options
        )
        print("Client initialized successfully")
        
        # Read PDF
        pdf_path = os.path.join(UPLOAD_FOLDER, 'offer1.pdf')
        print("Reading PDF: " + pdf_path)
        
        if not os.path.exists(pdf_path):
            print("ERROR: PDF not found")
            return False
        
        with open(pdf_path, 'rb') as f:
            content = f.read()
        
        print("PDF read: " + str(len(content)) + " bytes")
        
        # Build processor path
        name = "projects/" + PROJECT_ID + "/locations/" + LOCATION + "/processors/" + PROCESSOR_ID
        print("Processor path: " + name)
        
        # Create request
        raw_doc = documentai.RawDocument(content=content, mime_type='application/pdf')
        request = documentai.ProcessRequest(name=name, raw_document=raw_doc)
        
        # Process document
        print("Calling Document AI API...")
        result = client.process_document(request=request)
        text = result.document.text
        
        print("Extracted: " + str(len(text)) + " characters")
        
        # Save text
        output_path = os.path.join(OUTPUT_FOLDER, 'extracted_text.txt')
        print("Saving to: " + output_path)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text)
        
        # Verify
        if os.path.exists(output_path):
            size = os.path.getsize(output_path)
            print("SUCCESS: File saved")
            print("Size: " + str(size) + " bytes")
            print("Location: " + output_path)
            
            # List output folder
            files = os.listdir(OUTPUT_FOLDER)
            print("Output folder contents: " + str(files))
            
            return True
        else:
            print("ERROR: File not created")
            return False
            
    except Exception as e:
        print("ERROR: " + str(e))
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("DOCUMENT AI PROCESSING")
    print("=" * 60)
    
    success = main()
    
    if success:
        print("=" * 60)
        print("COMPLETED SUCCESSFULLY")
        print("=" * 60)
        sys.exit(0)
    else:
        print("=" * 60)
        print("FAILED")
        print("=" * 60)
        sys.exit(1)