import sys
print("Python version:", sys.version, flush=True)

try:
    import fitz
    print("PyMuPDF imported successfully", flush=True)
    print("PyMuPDF version:", fitz.version, flush=True)
except Exception as e:
    print(f"PyMuPDF import failed: {e}", flush=True)
    sys.exit(1)

try:
    import openai
    print("OpenAI imported successfully", flush=True)
except Exception as e:
    print(f"OpenAI import failed: {e}", flush=True)
    sys.exit(1)

print("All imports OK", flush=True)
sys.exit(0)