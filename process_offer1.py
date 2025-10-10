from google.cloud import documentai_v1 as documentai
import os
import json
import tempfile

def process_offer1(file_path):
    """
    Process Offer 1 (supplier quotation) using Google Document AI
    
    Args:
        file_path: Path to the PDF file to process
        
    Returns:
        Document object with extracted text and structure
    """
    
    # === STEP 1: Setup Google Cloud Credentials ===
    # Check if credentials are provided as environment variable (for Render)
    creds_json = os.getenv('GOOGLE_APPLICATION_CREDENTIALS_JSON')
    
    if creds_json:
        # Running on Render - create temporary credentials file
        try:
            creds_dict = json.loads(creds_json)
            with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.json') as temp_creds:
                json.dump(creds_dict, temp_creds)
                temp_creds_path = temp_creds.name
            
            # Set the path for Google Cloud to find credentials
            os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = temp_creds_path
            print(f"✅ Using credentials from environment variable")
        except Exception as e:
            print(f"⚠️  Error setting up credentials: {e}")
            # Try to continue anyway - might work with default credentials
    else:
        # Running locally - use file path
        print("ℹ️  Using local credentials file")
        # Credentials should be set via GOOGLE_APPLICATION_CREDENTIALS env var locally
    
    # === STEP 2: Configure Document AI ===
    project_id = os.getenv("GCP_PROJECT_ID", "requote-ai-backend")
    location = os.getenv("GCP_LOCATION", "eu")  # Correct: your processor is in EU
    processor_id = os.getenv("GCP_PROCESSOR_ID", "f02a4802c23ab664")
    mime_type = "application/pdf"
    
    print(f"📄 Processing document with Document AI...")
    print(f"   Project: {project_id}")
    print(f"   Location: {location}")
    print(f"   Processor: {processor_id}")
    
    # === STEP 3: Create Document AI Client ===
    from google.api_core.client_options import ClientOptions
    
    # Set the regional endpoint
    client_options = ClientOptions(
        api_endpoint=f"{location}-documentai.googleapis.com"
    )
    
    try:
        client = documentai.DocumentProcessorServiceClient(client_options=client_options)
        print(f"✅ Document AI client created successfully")
    except Exception as e:
        print(f"❌ Failed to create Document AI client: {e}")
        raise
    
    # === STEP 4: Build Processor Name ===
    name = f"projects/{project_id}/locations/{location}/processors/{processor_id}"
    print(f"📍 Processor path: {name}")
    
    # === STEP 5: Read File and Process ===
    try:
        with open(file_path, "rb") as file:
            document = {"content": file.read(), "mime_type": mime_type}
        
        print(f"📤 Sending document to Document AI...")
        
        # Create the request
        request = {"name": name, "raw_document": document}
        
        # Process the document
        result = client.process_document(request=request)
        
        print(f"✅ Document processed successfully!")
        print(f"   Pages: {len(result.document.pages)}")
        print(f"   Text length: {len(result.document.text)} characters")
        
        return result.document
        
    except FileNotFoundError:
        print(f"❌ File not found: {file_path}")
        raise
    except Exception as e:
        print(f"❌ Error processing document: {e}")
        raise
    finally:
        # Clean up temporary credentials file if it exists
        if creds_json and 'temp_creds_path' in locals():
            try:
                os.unlink(temp_creds_path)
                print(f"🧹 Cleaned up temporary credentials file")
            except:
                pass

# Main execution for testing
if __name__ == "__main__":
    # For local testing
    file_path = "uploads/offer1.pdf"
    
    if os.path.exists(file_path):
        print(f"🚀 Starting Document AI processing...")
        try:
            doc_result = process_offer1(file_path)
            parsed_text = doc_result.text
            
            # Save to output
            output_path = "outputs/parsed_offer1.txt"
            os.makedirs("outputs", exist_ok=True)
            
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(parsed_text)
            
            print(f"✅ Parsed text saved to: {output_path}")
            print(f"📊 Preview (first 500 chars):")
            print(parsed_text[:500])
            
        except Exception as e:
            print(f"❌ Processing failed: {e}")
            import traceback
            traceback.print_exc()
    else:
        print(f"❌ File not found: {file_path}")
        print(f"ℹ️  Please upload a PDF file to the uploads folder first")