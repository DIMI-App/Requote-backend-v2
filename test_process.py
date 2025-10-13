import os
import sys

from process_offer1 import process_offer1_and_save

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")


def main() -> int:
    file_path = os.path.join(UPLOAD_FOLDER, "offer1.pdf")
    output_path = os.path.join(OUTPUT_FOLDER, "extracted_text.txt")

    if not os.path.exists(file_path):
        print(f"❌ File not found: {file_path}")
        print("ℹ️  Please upload a PDF file to the uploads folder first")
        return 1

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    success = process_offer1_and_save(file_path, output_path)
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
