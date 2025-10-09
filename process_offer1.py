from google.cloud import documentai_v1 as documentai
import os

def process_offer1(file_path):
    project_id = "requote-ai-backend"
    location = "eu"  # Correct: your processor is in EU
    processor_id = "f02a4802c23ab664"
    mime_type = "application/pdf"

    # Set the regional endpoint
    client_options = {"api_endpoint": f"{location}-documentai.googleapis.com"}

    # Create Document AI client with EU endpoint
    client = documentai.DocumentProcessorServiceClient(client_options=client_options)

    name = f"projects/{project_id}/locations/{location}/processors/{processor_id}"

    with open(file_path, "rb") as file:
        document = {"content": file.read(), "mime_type": mime_type}

    request = {"name": name, "raw_document": document}
    result = client.process_document(request=request)

    return result.document
