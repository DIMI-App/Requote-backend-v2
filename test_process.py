import os
import sys
import json
import tempfile
from google.cloud import documentai_v1 as documentai
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
    temp_path = None
    
    print("Starting Document AI processing...")
    
    if not PROJECT_ID or not PROCESSOR_ID or not CREDENTIALS_JSON:
        print("ERROR: Missing environment variables")
        return False
    
    print("Project ID: " + PROJECT_ID)
    print("Processor ID: " + PROCESSOR_ID)
    
    try:
        creds_dict = json.loads(CREDENTIALS_JSON)
        print("Credentials loaded")
        
        temp_file = tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.json')
        json.dump(creds_dict, temp_file)
        temp_file.close()
        temp_path = temp_file.name
        
        os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = temp_path
        print("Temp credentials created")
        
        client_options = ClientOptions(api_endpoint=LOCATION + "-documentai.googleapis.com")
        client = documentai.DocumentProcessorServiceClient(client_options=client_options)
        print("Client initialized")
        
        pdf_path = os.path.join(UPLOAD_FOLDER, 'offer1.pdf')
        print("Reading PDF: " + pdf_path)
        
        with open(pdf_path, 'rb') as f:
            content = f.read()
        
        print("PDF read: " + str(len(content)) + " bytes")
        
        name = "projects/" + PROJECT_ID + "/locations/" + LOCATION + "/processors/" + PROCESSOR_ID
        
        raw_doc = documentai.RawDocument(content=content, mime_type='application/pdf')
        request = documentai.ProcessRequest(name=name, raw_document=raw_doc)
        
        print("Processing document...")
        result = client.process_document(request=request)
        text = result.document.text
        
        print("Extracted: " + str(len(text)) + " characters")
        
        output_path = os.path.join(OUTPUT_FOLDER, 'extracted_text.txt')
        print("Saving to: " + output_path)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text)
        
        if os.path.exists(output_path):
            size = os.path.getsize(output_path)
            print("SUCCESS: File saved (" + str(size) + " bytes)")
            print("Location: " + output_path)
            return True
        else:
            print("ERROR: File not created")
            return False
            
    except Exception as e:
        print("ERROR: " + str(e))
        import traceback
        traceback.print_exc()
        return False
    finally:
        if temp_path and os.path.exists(temp_path):
            os.unlink(temp_path)
            print("Cleaned up temp file")

if __name__ == "__main__":
    success = main()
    if success:
        print("COMPLETED")
        sys.exit(0)
    else:
        print("FAILED")
        sys.exit(1)