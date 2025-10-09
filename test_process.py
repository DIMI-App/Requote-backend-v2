import os
from process_offer1 import process_offer1
from google.cloud.documentai_v1.types.document import Document

# Set the path to your Google Cloud service account credentials
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "C:\\Users\\vital\\Desktop\\Requote-backend\\requote-ai-backend-12ffe3b92829.json"

# Path to the file you want to process
file_path = "uploads/offer1.pdf"

# Call the processing function
document: Document = process_offer1(file_path)

# Extract plain text
parsed_text = document.text

# Save to a .txt file
output_path = "outputs/parsed_offer1.txt"
os.makedirs("outputs", exist_ok=True)
with open(output_path, "w", encoding="utf-8") as f:
    f.write(parsed_text)

print(f"✅ Parsed text saved to: {output_path}")
