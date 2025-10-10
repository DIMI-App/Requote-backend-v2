from google.cloud import documentai_v1 as documentai
import os
import json
import tempfile
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def process_and_save():
    """Process PDF and save extracted text"""
    
    # Setup Google Cloud Credentials
    creds_json = os.getenv('GOOGLE_APPLICATION_CREDENTIALS_JSON')
    temp_creds_path = None
    
    if creds_json:
        try:
            creds_dict = json.loads(creds_json)
            with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.json') as temp_creds:
                json.dump(creds_dict, temp_creds)
                temp_creds_path = temp_creds.name
            
            os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = temp_creds_path
            print("Using credentials from environment variable")
        except Exception as e:
            print("Error setting up credentials: " + str(e))
            return False
    
    # Configure Document AI
    project_id = os.getenv("GCP_PROJECT_ID")
    location = "us"
    processor_id = os.getenv("GCP_PROCESSOR_ID")
    mime_type = "application/pdf"
    
    print("Processing document with Document AI...")
    print("Project: " + project_id)
    print("Processor: " + processor_id)
    
    # Create Document AI Client
    from google.api_core.client_options import ClientOptions
    
    client_options = ClientOptions(
        api_endpoint=location + "-documentai.googleapis.com"
    )
    
    try:
        client = documentai.DocumentProcessorServiceClient(client_options=client_options)
        print("Document AI client created successfully")
    except Exception as e:
        print("Failed to create Document AI client: " + str(e))
        return False
    
    # Build Processor Name
    name = "projects/" + project_id + "/locations/" + location + "/processors/" + processor_id
    
    # Read and Process File
    file_path = os.path.join(UPLOAD_FOLDER, "offer1.pdf")
    
    try:
        with open(file_path, "rb") as file:
            document = {"content": file.read(), "mime_type": mime_type}
        
        print("Sending document to Document AI...")
        
        request = {"name": name, "raw_document": document}
        result = client.process_document(request=request)
        
        extracted_text = result.document.text
        
        print("Document processed successfully!")
        print("Text length: " + str(len(extracted_text)) + " characters")
        
        # Save extracted text
        output_path = os.path.join(OUTPUT_FOLDER, "extracted_text.txt")
        print("Saving to: " + output_path)
        
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(extracted_text)
        
        # Verify file was saved
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print("SUCCESS: File saved (" + str(file_size) + " bytes)")
            print("Location: " + output_path)
            
            # List files in output folder
            files = os.listdir(OUTPUT_FOLDER)
            print("Files in output folder: " + str(files))
            
            return True
        else:
            print("ERROR: File was not created")
            return False
        
    except Exception as e:
        print("Error: " + str(e))
        import traceback
        traceback.print_exc()
        return False
    
    finally:
        # Clean up temporary credentials file
        if temp_creds_path:
            try:
                os.unlink(temp_creds_path)
                print("Cleaned up temporary credentials file")
            except:
                pass

if __name__ == "__main__":
    print("=" * 60)
    print("STARTING DOCUMENT AI PROCESSING")
    print("=" * 60)
    
    success = process_and_save()
    
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