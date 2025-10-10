from google.cloud import documentai_v1 as documentai
import os
import json
import tempfile

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
    
    if creds_json:
        try:
            creds_dict = json.loads(creds_json)
            with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.json') as temp_creds:
                json.dump(creds_dict, temp_creds)
                temp_creds_path = temp_creds.name
            
            os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = temp_creds_path
            print("✅ Using credentials from environment variable")
        except Exception as e:
            print("⚠️  Error setting up credentials: " + str(e))
    
    # === STEP 2: Configure Document AI ===
    project_id = os.getenv("GCP_PROJECT_ID", "requote-ai-backend")
    location = os.getenv("GCP_LOCATION", "eu")
    processor_id = os.getenv("GCP_PROCESSOR_ID", "f02a4802c23ab664")
    mime_type = "application/pdf"
    
    print("📄 Processing document with Document AI...")
    print("   Project: " + project_id)
    print("   Location: " + location)
    print("   Processor: " + processor_id)
    
    # === STEP 3: Create Document AI Client ===
    from google.api_core.client_options import ClientOptions
    
    client_options = ClientOptions(
        api_endpoint=location + "-documentai.googleapis.com"
    )
    
    try:
        client = documentai.DocumentProcessorServiceClient(client_options=client_options)
        print("✅ Document AI client created successfully")
    except Exception as e:
        print("❌ Failed to create Document AI client: " + str(e))
        raise
    
    # === STEP 4: Build Processor Name ===
    name = "projects/" + project_id + "/locations/" + location + "/processors/" + processor_id
    print("📍 Processor path: " + name)
    
    # === STEP 5: Read File and Process ===
    try:
        with open(file_path, "rb") as file:
            document = {"content": file.read(), "mime_type": mime_type}
        
        print("📤 Sending document to Document AI...")
        
        request = {"name": name, "raw_document": document}
        result = client.process_document(request=request)
        
        print("✅ Document processed successfully!")
        print("   Pages: " + str(len(result.document.pages)))
        print("   Text length: " + str(len(result.document.text)) + " characters")
        
        return result.document
        
    except FileNotFoundError:
        print("❌ File not found: " + file_path)
        raise
    except Exception as e:
        print("❌ Error processing document: " + str(e))
        raise
    finally:
        if creds_json and 'temp_creds_path' in locals():
            try:
                os.unlink(temp_creds_path)
                print("🧹 Cleaned up temporary credentials file")
            except:
                pass

if __name__ == "__main__":
    file_path = os.path.join(UPLOAD_FOLDER, "offer1.pdf")
    
    if os.path.exists(file_path):
        print("🚀 Starting Document AI processing...")
        try:
            doc_result = process_offer1(file_path)
            parsed_text = doc_result.text
            
            # === THE ONLY CHANGE: Save to extracted_text.txt instead of parsed_offer1.txt ===
            output_path = os.path.join(OUTPUT_FOLDER, "extracted_text.txt")
            os.makedirs(OUTPUT_FOLDER, exist_ok=True)
            
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(parsed_text)
            
            print("✅ Parsed text saved to: " + output_path)
            print("📊 Preview (first 500 chars):")
            print(parsed_text[:500])
            
        except Exception as e:
            print("❌ Processing failed: " + str(e))
            import traceback
            traceback.print_exc()
    else:
        print("❌ File not found: " + file_path)
        print("ℹ️  Please upload a PDF file to the uploads folder first")