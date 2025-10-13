import base64
import json
import os
import tempfile
from typing import Any, Dict, Optional, Tuple

from google.api_core import exceptions as google_exceptions
from google.api_core.client_options import ClientOptions
from google.cloud import documentai_v1 as documentai
from pdfminer.high_level import extract_text as pdfminer_extract_text
from pypdf import PdfReader

DEFAULT_TIMEOUT_SECONDS = int(os.getenv("DOCUMENT_AI_TIMEOUT", "110"))


class DocumentAICredentialsError(RuntimeError):
    """Raised when Document AI credentials are missing or invalid."""


def _maybe_decode_credentials(raw_value: str) -> Dict[str, Any]:
    """Attempt to interpret the provided credential payload as JSON or base64."""

    try:
        return json.loads(raw_value)
    except json.JSONDecodeError:
        pass

    try:
        decoded_bytes = base64.b64decode(raw_value)
        decoded = decoded_bytes.decode("utf-8")
    except Exception as exc:  # pragma: no cover - defensive
        raise DocumentAICredentialsError(
            "GOOGLE_APPLICATION_CREDENTIALS_JSON is not valid JSON or base64-encoded JSON"
        ) from exc

    try:
        return json.loads(decoded)
    except json.JSONDecodeError as exc:  # pragma: no cover - defensive
        raise DocumentAICredentialsError(
            "Decoded GOOGLE_APPLICATION_CREDENTIALS_JSON is not valid JSON"
        ) from exc


def _setup_credentials() -> Optional[str]:
    """Prepare Google Cloud credentials and return the temp file path if used."""

    if os.getenv("DISABLE_DOCUMENT_AI", "").lower() in {"1", "true", "yes"}:
        raise DocumentAICredentialsError(
            "Document AI disabled via DISABLE_DOCUMENT_AI environment variable"
        )

    creds_json = os.getenv("GOOGLE_APPLICATION_CREDENTIALS_JSON")

    if creds_json:
        creds_json = creds_json.strip()
        creds_dict = _maybe_decode_credentials(creds_json)

        with tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".json") as temp_creds:
            json.dump(creds_dict, temp_creds)
            temp_creds_path = temp_creds.name

        os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = temp_creds_path
        print("✅ Using credentials provided via GOOGLE_APPLICATION_CREDENTIALS_JSON")
        return temp_creds_path

    creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if creds_path:
        if os.path.exists(creds_path):
            print("✅ Using credentials from GOOGLE_APPLICATION_CREDENTIALS path")
            return None

        raise DocumentAICredentialsError(
            "GOOGLE_APPLICATION_CREDENTIALS points to a missing file"
        )

    raise DocumentAICredentialsError(
        "No Document AI credentials configured. Set GOOGLE_APPLICATION_CREDENTIALS_JSON or GOOGLE_APPLICATION_CREDENTIALS."
    )


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


def _extract_text_with_pypdf(file_path: str) -> str:
    """Fast fallback using PyPDF to keep response times low."""
    print("⚠️  Falling back to PyPDF text extraction")
    try:
        reader = PdfReader(file_path)
        pieces = []
        for page_number, page in enumerate(reader.pages, start=1):
            try:
                pieces.append(page.extract_text() or "")
            except Exception as exc:  # pragma: no cover - defensive
                print(f"   ⚠️  PyPDF failed on page {page_number}: {exc}")
        text = "\n".join(filter(None, pieces))
        print(f"✅ PyPDF extracted {len(text)} characters")
        return text
    except Exception as exc:  # pragma: no cover - dependency
        print(f"❌ PyPDF extraction failed: {exc}")
        return ""


def _fallback_extract_text(file_path: str) -> str:
    """Extract text locally using progressively heavier fallbacks."""
    text = _extract_text_with_pypdf(file_path)
    if len(text) >= 500:  # Heuristic: consider PyPDF result good enough
        return text

    if text:
        print("⚠️  PyPDF produced very little text, trying pdfminer...")
    else:
        print("⚠️  PyPDF returned no text, trying pdfminer...")

    try:
        text = pdfminer_extract_text(file_path)
        print(f"✅ pdfminer extracted {len(text)} characters")
        return text
    except Exception as exc:  # pragma: no cover - dependency
        print(f"❌ pdfminer extraction failed: {exc}")
        return ""


def _describe_document_ai_error(error_type: str, exc: Exception) -> Dict[str, Any]:
    status = getattr(exc, "code", None)
    if status is not None:
        status = str(status)
    return {
        "type": error_type,
        "message": str(exc),
        "details": getattr(exc, "errors", None),
        "status": status,
    }


def extract_offer1_text(
    file_path: str, timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS
) -> Tuple[str, Dict[str, Any]]:
    """Attempt Document AI extraction and fall back to local parsing on failure.

    Returns a tuple of the extracted text and a diagnostics dictionary describing
    which extractor produced the text and any errors encountered along the way.
    """
    diagnostics: Dict[str, Any] = {
        "document_ai_status": "not_attempted",
        "used_fallback": False,
    }
    document_ai_error: Optional[Dict[str, Any]] = None
    try:
        document = process_offer1(file_path, timeout_seconds=timeout_seconds)
        text = getattr(document, "text", "") or ""
        if text.strip():
            print("✅ Using text extracted by Document AI")
            diagnostics["document_ai_status"] = "success"
            diagnostics["document_ai_characters"] = len(text)
            return text, diagnostics

        print("⚠️  Document AI returned empty text. Using fallback extractor.")
        document_ai_error = {
            "type": "document_ai_empty_text",
            "message": "Document AI returned no text",
        }
    except DocumentAICredentialsError as exc:
        print(f"🚫 Document AI skipped: {exc}")
        diagnostics["document_ai_status"] = "skipped"
        document_ai_error = {
            "type": "document_ai_credentials",
            "message": str(exc),
        }
    except google_exceptions.DeadlineExceeded as exc:
        print("⏱️  Document AI request exceeded timeout. Using fallback extractor.")
        document_ai_error = _describe_document_ai_error("document_ai_timeout", exc)
    except google_exceptions.PermissionDenied as exc:
        print("🚫 Document AI credentials lack required permissions. Using fallback extractor.")
        document_ai_error = _describe_document_ai_error("document_ai_permission", exc)
    except google_exceptions.Unauthenticated as exc:
        print("🚫 Document AI authentication failed. Using fallback extractor.")
        document_ai_error = _describe_document_ai_error("document_ai_unauthenticated", exc)
    except Exception as exc:  # pragma: no cover - network dependency
        print(f"❌ Document AI processing failed: {exc}")
        document_ai_error = _describe_document_ai_error("document_ai_error", exc)

    if diagnostics["document_ai_status"] == "not_attempted":
        diagnostics["document_ai_status"] = "failed"

    if document_ai_error:
        diagnostics["document_ai_error"] = document_ai_error

    fallback_text = _fallback_extract_text(file_path)
    diagnostics["used_fallback"] = True
    diagnostics["fallback_characters"] = len(fallback_text)
    return fallback_text, diagnostics


def save_text_to_file(text: str, output_path: str) -> None:
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as file:
        file.write(text)
    print(f"💾 Saved extracted text to {output_path}")


def process_offer1_and_save(
    file_path: str,
    output_path: str,
    timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS,
) -> bool:
    text, diagnostics = extract_offer1_text(
        file_path, timeout_seconds=timeout_seconds
    )
    if not text.strip():
        print("❌ No text extracted from Offer 1")
        if diagnostics.get("document_ai_error"):
            print(
                "   Document AI error:",
                diagnostics["document_ai_error"].get("message", "unknown"),
            )
        return False

    save_text_to_file(text, output_path)
    preview = text[:500].replace("\n", " ")
    print(" Preview (first 500 chars):")
    print(preview)
    if diagnostics.get("used_fallback"):
        print("ℹ️  Text was generated using local fallback extraction")
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
