from google.cloud import documentai_v1 as documentai
import os
import json
import tempfile
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def process_offer1(file_path):
    """
    Process Offer 1 (supplier quotation) using Google Document AI
    """
    
    # === STEP 1: Setup Google Cloud Credentials ===
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
    
    # === STEP 2: Configure Document AI ===
    project_id = os.getenv("GCP_PROJECT_ID")
    location = "us"
    processor_id = os.getenv("GCP_PROCESSOR_ID")
    mime_type = "application/pdf"
    
    print("Processing document with Document AI...")
    print("Project: " + project_id)
    print("Location: " + location)
    print("Processor: " + processor_id)
    
    # === STEP 3: Create Document AI Client ===
    from google.api_core.client_options import ClientOptions
    
    client_options = ClientOptions(
        api_endpoint=location + "-documentai.googleapis.com"
    )
    
    try:
        client = documentai.DocumentProcessorServiceClient(client_options=client_options)
        print("Document AI client created successfully")
    except Exception as e:
        print("Failed to create Document AI client: " + str(e))
        return None, temp_creds_path
    
    # === STEP 4: Build Processor Name ===
    name = "projects/" + project_id + "/locations/" + location + "/processors/" + processor_id
    print("Processor path: " + name)
    
    # === STEP 5: Read File and Process ===
    try:
        with open(file_path, "rb") as file:
            document = {"content": file.read(), "mime_type": mime_type}
        
        print("Sending document to Document AI...")
        
        request = {"name": name, "raw_document": document}
        result = client.process_document(request=request)
        
        print("Document processed successfully!")
        print("Pages: " + str(len(result.document.pages)))
        print("Text length: " + str(len(result.document.text)) + " characters")
        
        return result.document, temp_creds_path
        
    except Exception as e:
        print("Error processing document: " + str(e))
        import traceback
        traceback.print_exc()
        return None, temp_creds_path

def main():
    """Main execution function"""
    print("=" * 60)
    print("STARTING DOCUMENT AI PROCESSING")
    print("=" * 60)
    
    file_path = os.path.join(UPLOAD_FOLDER, "offer1.pdf")
    
    if not os.path.exists(file_path):
        print("ERROR: File not found: " + file_path)
        return False
    
    print("Input file: " + file_path)
    
    # Process document
    doc_result, temp_creds_path = process_offer1(file_path)
    
    # Clean up credentials
    if temp_creds_path:
        try:
            os.unlink(temp_creds_path)
            print("Cleaned up temporary credentials file")
        except:
            pass
    
    if not doc_result:
        print("ERROR: Processing failed")
        return False
    
    # Extract text
    parsed_text = doc_result.text
    
    # Save to the correct filename
    output_path = os.path.join(OUTPUT_FOLDER, "extracted_text.txt")
    
    print("Saving extracted text to: " + output_path)
    
    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(parsed_text)
        
        print("File write completed")
        
        # Verify the file was created
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print("SUCCESS: File saved!")
            print("Location: " + output_path)
            print("Size: " + str(file_size) + " bytes")
            
            # Show what's in the output folder
            files_in_output = os.listdir(OUTPUT_FOLDER)
            print("Files in output folder: " + str(files_in_output))
            
            return True
        else:
            print("ERROR: File was not created")
            print("Output folder exists: " + str(os.path.exists(OUTPUT_FOLDER)))
            print("Files in output folder: " + str(os.listdir(OUTPUT_FOLDER)))
            return False
            
    except Exception as e:
        print("ERROR: Failed to save file: " + str(e))
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
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