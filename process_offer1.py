import json
import os
import tempfile
from typing import Optional

from google.api_core import exceptions as google_exceptions
from google.api_core.client_options import ClientOptions
from google.cloud import documentai_v1 as documentai
from pdfminer.high_level import extract_text as pdfminer_extract_text

DEFAULT_TIMEOUT_SECONDS = int(os.getenv("DOCUMENT_AI_TIMEOUT", "65"))


def _setup_credentials() -> Optional[str]:
    """Prepare Google Cloud credentials and return the temp file path if used."""
    creds_json = os.getenv("GOOGLE_APPLICATION_CREDENTIALS_JSON")

    if not creds_json:
        print("ℹ️  Using GOOGLE_APPLICATION_CREDENTIALS file path from environment")
        return None

    try:
        creds_dict = json.loads(creds_json)
        with tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".json") as temp_creds:
            json.dump(creds_dict, temp_creds)
            temp_creds_path = temp_creds.name

        os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = temp_creds_path
        print("✅ Using credentials provided via GOOGLE_APPLICATION_CREDENTIALS_JSON")
        return temp_creds_path
    except Exception as exc:  # pragma: no cover - defensive
        print(f"⚠️  Failed to configure credentials from JSON: {exc}")
        return None


def process_offer1(file_path: str, timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS):
    """Run Document AI on the supplied PDF and return the processed document."""
    temp_creds_path = _setup_credentials()

    project_id = os.getenv("GCP_PROJECT_ID", "requote-ai-backend")
    location = os.getenv("GCP_LOCATION", "eu")
    processor_id = os.getenv("GCP_PROCESSOR_ID", "f02a4802c23ab664")

    print("📄 Processing document with Document AI...")
    print(f"   Project: {project_id}")
    print(f"   Location: {location}")
    print(f"   Processor: {processor_id}")

    client_options = ClientOptions(api_endpoint=f"{location}-documentai.googleapis.com")

    try:
        client = documentai.DocumentProcessorServiceClient(client_options=client_options)
        print("✅ Document AI client created successfully")
    except Exception as exc:  # pragma: no cover - network dependency
        print(f"❌ Failed to create Document AI client: {exc}")
        raise

    name = f"projects/{project_id}/locations/{location}/processors/{processor_id}"
    print(f"📍 Processor path: {name}")

    try:
        with open(file_path, "rb") as file:
            raw_document = {"content": file.read(), "mime_type": "application/pdf"}

        request = {"name": name, "raw_document": raw_document}
        result = client.process_document(request=request, timeout=timeout_seconds)

        print("✅ Document processed successfully!")
        print(f"   Pages: {len(result.document.pages)}")
        print(f"   Text length: {len(result.document.text)} characters")
        return result.document
    except FileNotFoundError:
        print(f"❌ File not found: {file_path}")
        raise
    except Exception as exc:  # pragma: no cover - network dependency
        print(f"❌ Error processing document with Document AI: {exc}")
        raise
    finally:
        if temp_creds_path and os.path.exists(temp_creds_path):
            try:
                os.unlink(temp_creds_path)
                print("🧹 Cleaned up temporary credentials file")
            except OSError:
                pass


def _fallback_extract_text(file_path: str) -> str:
    """Extract text locally from the PDF using pdfminer as a fallback."""
    print("⚠️  Falling back to local PDF text extraction (pdfminer)")
    try:
        text = pdfminer_extract_text(file_path)
        print(f"✅ Fallback extracted {len(text)} characters")
        return text
    except Exception as exc:  # pragma: no cover - dependency
        print(f"❌ Fallback extraction failed: {exc}")
        return ""


def extract_offer1_text(file_path: str, timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> str:
    """Attempt Document AI extraction and fall back to local parsing on failure."""
    try:
        document = process_offer1(file_path, timeout_seconds=timeout_seconds)
        text = getattr(document, "text", "") or ""
        if text.strip():
            print("✅ Using text extracted by Document AI")
            return text
        print("⚠️  Document AI returned empty text. Using fallback extractor.")
    except google_exceptions.DeadlineExceeded:
        print("⏱️  Document AI request exceeded timeout. Using fallback extractor.")
    except Exception as exc:  # pragma: no cover - network dependency
        print(f"❌ Document AI processing failed: {exc}")

    return _fallback_extract_text(file_path)


def save_text_to_file(text: str, output_path: str) -> None:
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as file:
        file.write(text)
    print(f"💾 Saved extracted text to {output_path}")


def process_offer1_and_save(file_path: str, output_path: str, timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> bool:
    text = extract_offer1_text(file_path, timeout_seconds=timeout_seconds)
    if not text.strip():
        print("❌ No text extracted from Offer 1")
        return False

    save_text_to_file(text, output_path)
    preview = text[:500].replace("\n", " ")
    print(" Preview (first 500 chars):")
    print(preview)
    return True


if __name__ == "__main__":
    FILE_PATH = os.path.join("uploads", "offer1.pdf")
    OUTPUT_PATH = os.path.join("outputs", "extracted_text.txt")

    if not os.path.exists(FILE_PATH):
        print(f"❌ File not found: {FILE_PATH}")
        print("ℹ️  Please upload a PDF file to the uploads folder first")
        raise SystemExit(1)

    success = process_offer1_and_save(FILE_PATH, OUTPUT_PATH)
    raise SystemExit(0 if success else 1)
