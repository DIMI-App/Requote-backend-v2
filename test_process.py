from google.cloud import documentai_v1 as documentai
import os
import json
import tempfile
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def diagnose_credentials():
    """Check if credentials are valid"""
    print("=" * 60)
    print("DIAGNOSING CREDENTIALS")
    print("=" * 60)
    
    creds_json = os.getenv('GOOGLE_APPLICATION_CREDENTIALS_JSON')
    
    if not creds_json:
        print("ERROR: GOOGLE_APPLICATION_CREDENTIALS_JSON not set")
        return None
    
    print("Credentials found in environment")
    print("Length: " + str(len(creds_json)) + " characters")
    
    try:
        creds_dict = json.loads(creds_json)
        print("JSON parsed successfully")
        
        # Check for required keys
        required_keys = ['type', 'project_id', 'private_key', 'client_email']
        for key in required_keys:
            if key in creds_dict:
                if key == 'private_key':
                    print("✓ " + key + ": [PRESENT - " + str(len(creds_dict[key])) + " chars]")
                else:
                    print("✓ " + key + ": " + str(creds_dict[key]))
            else:
                print("✗ " + key + ": MISSING")
                return None
        
        return creds_dict
        
    except json.JSONDecodeError as e:
        print("ERROR: Invalid JSON - " + str(e))
        print("First 200 chars: " + creds_json[:200])
        return None

def process_and_save():
    """Process PDF and save extracted text"""
    
    # Diagnose credentials first
    creds_dict = diagnose_credentials()
    if not creds_dict:
        print("FATAL: Invalid credentials")
        return False
    
    print("")
    print("=" * 60)
    print("PROCESSING DOCUMENT")
    print("=" * 60)
    
    temp_creds_path = None
    
    try:
        # Create temp credentials file
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.json') as temp_creds:
            json.dump(creds_dict, temp_creds)
            temp_creds_path = temp_creds.name
        
        os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = temp_creds_path
        print("Temp credentials file: " + temp_creds_path)
        
        # Verify temp file was created
        if os.path.exists(temp_creds_path):
            temp_size = os.path.getsize(temp_creds_path)
            print("Temp file verified: " + str(temp_size) + " bytes")
        else:
            print("ERROR: Temp file not created")
            return False
        
        # Configure Document AI
        project_id = os.getenv("GCP_PROJECT_ID")
        location = "us"
        processor_id = os.getenv("GCP_PROCESSOR_ID")
        
        print("Project: " + project_id)
        print("Processor: " + processor_id)
        
        # Create client
        from google.api_core.client_options import ClientOptions
        
        client_options = ClientOptions(
            api_endpoint=location + "-documentai.googleapis.com"
        )
        
        print("Creating Document AI client...")
        client = documentai.DocumentProcessorServiceClient(client_options=client_options)
        print("Client created")
        
        # Read PDF
        file_path = os.path.join(UPLOAD_FOLDER, "offer1.pdf")
        print("Reading: " + file_path)
        
        with open(file_path, "rb") as file:
            content = file.read()
        
        print("PDF loaded: " + str(len(content)) + " bytes")
        
        # Process
        name = "projects/" + project_id + "/locations/" + location + "/processors/" + processor_id
        
        document = {"content": content, "mime_type": "application/pdf"}
        request = {"name": name, "raw_document": document}
        
        print("Calling Document AI...")
        result = client.process_document(request=request)
        
        extracted_text = result.document.text
        print("Extracted: " + str(len(extracted_text)) + " characters")
        
        # Save
        output_path = os.path.join(OUTPUT_FOLDER, "extracted_text.txt")
        
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(extracted_text)
        
        if os.path.exists(output_path):
            print("SUCCESS: File saved to " + output_path)
            return True
        else:
            print("ERROR: File not saved")
            return False
        
    except Exception as e:
        print("ERROR: " + str(e))
        import traceback
        traceback.print_exc()
        return False
    
    finally:
        if temp_creds_path:
            try:
                os.unlink(temp_creds_path)
                print("Cleaned up temp file")
            except:
                pass

if __name__ == "__main__":
    success = process_and_save()
    sys.exit(0 if success else 1)