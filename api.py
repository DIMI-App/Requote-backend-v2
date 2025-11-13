from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import json
import subprocess
from werkzeug.utils import secure_filename
import threading
import time
import shutil

app = Flask(__name__)

CORS(app, resources={
    r"/*": {
        "origins": "*",
        "methods": ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
        "allow_headers": ["Content-Type", "Authorization", "Accept"],
        "expose_headers": ["Content-Type"],
        "supports_credentials": False,
        "max_age": 3600
    }
})

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# Allowed file extensions
ALLOWED_OFFER1_EXTENSIONS = {'pdf', 'docx', 'doc', 'xlsx', 'xls', 'png', 'jpg', 'jpeg'}
ALLOWED_OFFER2_EXTENSIONS = {'docx', 'doc', 'xlsx', 'xls', 'pdf'}

# Thread-safe status storage
_status_lock = threading.Lock()
processing_status = {
    'status': 'idle',
    'message': '',
    'items_count': 0,
    'items': [],
    'started_at': None,
    'updated_at': None,
    'file_format': None
}

def allowed_file(filename, allowed_extensions):
    """Check if file extension is allowed"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in allowed_extensions

def get_file_extension(filename):
    """Get file extension in lowercase"""
    return filename.rsplit('.', 1)[1].lower() if '.' in filename else ''

def convert_to_pdf(input_path, output_path, file_format):
    """Convert various formats to PDF for unified processing"""
    try:
        print(f"Converting {file_format.upper()} to PDF...", flush=True)
        
        if file_format in ['docx', 'doc', 'xlsx', 'xls']:
            # Convert Office documents to PDF using LibreOffice
            result = subprocess.run(
                ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir',
                 os.path.dirname(output_path), input_path],
                capture_output=True,
                text=True,
                timeout=120
            )
            
            if result.returncode == 0:
                # LibreOffice outputs to same directory with .pdf extension
                converted_file = input_path.rsplit('.', 1)[0] + '.pdf'
                if os.path.exists(converted_file):
                    if converted_file != output_path:
                        shutil.move(converted_file, output_path)
                    print(f"✓ {file_format.upper()} converted to PDF", flush=True)
                    return True
            
            print(f"✗ Conversion failed: {result.stderr}", flush=True)
            return False
            
        elif file_format in ['png', 'jpg', 'jpeg']:
            # Images are supported directly by PyMuPDF
            print("✓ Image file - will be processed directly", flush=True)
            return True
            
        else:
            print(f"✗ Unsupported format: {file_format}", flush=True)
            return False
            
    except FileNotFoundError:
        print("✗ LibreOffice not installed, cannot convert documents", flush=True)
        return False
    except Exception as e:
        print(f"✗ Conversion error: {str(e)}", flush=True)
        return False

def convert_to_docx(input_path, output_path, file_format):
    """Convert various template formats to DOCX"""
    try:
        print(f"Converting {file_format.upper()} template to DOCX...", flush=True)
        
        if file_format == 'docx':
            # Already DOCX, just copy
            shutil.copy(input_path, output_path)
            print("✓ DOCX template ready", flush=True)
            return True
            
        elif file_format in ['doc', 'pdf', 'xlsx', 'xls']:
            # Convert to DOCX using LibreOffice
            result = subprocess.run(
                ['libreoffice', '--headless', '--convert-to', 'docx', '--outdir',
                 os.path.dirname(output_path), input_path],
                capture_output=True,
                text=True,
                timeout=120
            )
            
            if result.returncode == 0:
                converted_file = input_path.rsplit('.', 1)[0] + '.docx'
                if os.path.exists(converted_file):
                    if converted_file != output_path:
                        shutil.move(converted_file, output_path)
                    print(f"✓ {file_format.upper()} template converted to DOCX", flush=True)
                    return True
            
            print(f"✗ Template conversion failed: {result.stderr}", flush=True)
            return False
            
        else:
            print(f"✗ Unsupported template format: {file_format}", flush=True)
            return False
            
    except FileNotFoundError:
        print("✗ LibreOffice not installed, cannot convert template", flush=True)
        return False
    except Exception as e:
        print(f"✗ Template conversion error: {str(e)}", flush=True)
        return False

@app.after_request
def after_request(response):
    response.headers.add('Access-Control-Allow-Origin', '*')
    response.headers.add('Access-Control-Allow-Headers', 'Content-Type,Authorization')
    response.headers.add('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS')
    return response

@app.route('/', methods=['GET'])
def home():
    return jsonify({
        'message': 'Requote AI Backend is running!',
        'version': 'SV9-FullMultiFormat',
        'status': 'healthy',
        'supported_formats': {
            'offer1': list(ALLOWED_OFFER1_EXTENSIONS),
            'offer2': list(ALLOWED_OFFER2_EXTENSIONS)
        }
    })

def process_file_background(filepath, file_format):
    """Background processing for any supported file format"""
    global processing_status
    
    try:
        with _status_lock:
            processing_status['status'] = 'processing'
            processing_status['message'] = f'Processing {file_format.upper()} file...'
            processing_status['file_format'] = file_format
            processing_status['started_at'] = time.time()
            processing_status['updated_at'] = time.time()
        
        print("=== BACKGROUND PROCESSING STARTED ===", flush=True)
        print(f"Format: {file_format}", flush=True)
        print(f"File: {filepath}", flush=True)
        
        # Determine processing path based on format
        pdf_path = os.path.join(UPLOAD_FOLDER, 'offer1.pdf')
        
        if file_format == 'pdf':
            # Already PDF, just rename/copy
            if filepath != pdf_path:
                shutil.copy(filepath, pdf_path)
            print("✓ PDF ready for processing", flush=True)
            
        elif file_format in ['docx', 'doc', 'xlsx', 'xls']:
            # Convert to PDF first
            with _status_lock:
                processing_status['message'] = f'Converting {file_format.upper()} to PDF...'
                processing_status['updated_at'] = time.time()
            
            if not convert_to_pdf(filepath, pdf_path, file_format):
                with _status_lock:
                    processing_status['status'] = 'error'
                    processing_status['message'] = f'Failed to convert {file_format.upper()} to PDF. LibreOffice may not be installed.'
                    processing_status['updated_at'] = time.time()
                return
                
        elif file_format in ['png', 'jpg', 'jpeg']:
            # Images processed directly by extract script
            with _status_lock:
                processing_status['message'] = 'Processing image file...'
                processing_status['updated_at'] = time.time()
            
            # Copy image to expected location  
            shutil.copy(filepath, pdf_path.replace('.pdf', f'.{file_format}'))
            
        else:
            with _status_lock:
                processing_status['status'] = 'error'
                processing_status['message'] = f'Unsupported file format: {file_format}'
                processing_status['updated_at'] = time.time()
            return
        
        # Now run extraction
        with _status_lock:
            processing_status['message'] = 'Extracting items from document...'
            processing_status['updated_at'] = time.time()
        
        items_output_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        extract_script_path = os.path.join(BASE_DIR, 'extract_pdf_direct.py')
        
        print("Starting extraction subprocess...", flush=True)
        
        result = subprocess.run(
            ['python', extract_script_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=300
        )
        
        print(f"Extraction completed with code: {result.returncode}", flush=True)
        
        if result.stdout:
            print("=== EXTRACTION STDOUT ===", flush=True)
            print(result.stdout, flush=True)
        
        if result.stderr:
            print("=== EXTRACTION STDERR ===", flush=True)
            print(result.stderr, flush=True)
        
        if result.returncode != 0:
            with _status_lock:
                processing_status['status'] = 'error'
                processing_status['message'] = 'Extraction failed: ' + (result.stderr or result.stdout or 'Unknown error')
                processing_status['updated_at'] = time.time()
            return
        
        if not os.path.exists(items_output_path):
            with _status_lock:
                processing_status['status'] = 'error'
                processing_status['message'] = 'Items file not created after extraction'
                processing_status['updated_at'] = time.time()
            return
        
        with open(items_output_path, 'r', encoding='utf-8') as f:
            full_data = json.load(f)
        
        items = full_data.get('items', [])
        
        print(f"✓ Extracted {len(items)} items", flush=True)
        
        with _status_lock:
            processing_status['status'] = 'completed'
            processing_status['message'] = f'Successfully extracted {len(items)} items from {file_format.upper()}'
            processing_status['items_count'] = len(items)
            processing_status['items'] = items
            processing_status['updated_at'] = time.time()
        
        print("=== BACKGROUND PROCESSING COMPLETED ===", flush=True)
        elapsed = time.time() - processing_status['started_at']
        print(f"Total time: {elapsed:.1f} seconds", flush=True)
        
    except Exception as e:
        print(f"=== BACKGROUND PROCESSING ERROR: {str(e)} ===", flush=True)
        import traceback
        traceback.print_exc()
        
        with _status_lock:
            processing_status['status'] = 'error'
            processing_status['message'] = str(e)
            processing_status['updated_at'] = time.time()

@app.route('/api/process-offer1', methods=['POST', 'OPTIONS'])
def api_process_offer1():
    """Process Offer 1 - supports multiple formats"""
    global processing_status
    
    if request.method == 'OPTIONS':
        return '', 204
    
    try:
        print("=" * 60, flush=True)
        print("Received request to process Offer 1", flush=True)
        
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Validate file extension
        if not allowed_file(file.filename, ALLOWED_OFFER1_EXTENSIONS):
            return jsonify({
                'error': f'Unsupported file format. Allowed: {", ".join(ALLOWED_OFFER1_EXTENSIONS)}'
            }), 400
        
        file_extension = get_file_extension(file.filename)
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, f'offer1_original.{file_extension}')
        file.save(filepath)
        
        print(f"✓ File saved: {filepath}", flush=True)
        print(f"✓ Format: {file_extension.upper()}", flush=True)
        print(f"✓ Size: {os.path.getsize(filepath)} bytes", flush=True)
        
        # Reset status
        with _status_lock:
            processing_status = {
                'status': 'processing',
                'message': f'File uploaded ({file_extension.upper()}), starting processing...',
                'items_count': 0,
                'items': [],
                'file_format': file_extension,
                'started_at': time.time(),
                'updated_at': time.time()
            }
        
        # Start background thread
        thread = threading.Thread(
            target=process_file_background, 
            args=(filepath, file_extension), 
            daemon=True
        )
        thread.start()
        
        print(f"✓ Background thread started", flush=True)
        print("=" * 60, flush=True)
        
        return jsonify({
            'success': True,
            'message': f'Processing {file_extension.upper()} file. Poll /api/status for updates.',
            'status': 'processing',
            'file_format': file_extension
        })
        
    except Exception as e:
        print(f"ERROR in process-offer1: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/status', methods=['GET', 'OPTIONS'])
def api_status():
    if request.method == 'OPTIONS':
        return '', 204
    
    with _status_lock:
        status_copy = processing_status.copy()
    
    # Add elapsed time if processing
    if status_copy.get('started_at') and status_copy['status'] == 'processing':
        elapsed = time.time() - status_copy['started_at']
        status_copy['elapsed_seconds'] = round(elapsed, 1)
    
    return jsonify(status_copy)

@app.route('/api/upload-offer2', methods=['POST', 'OPTIONS'])
def api_upload_offer2():
    """Upload Offer 2 template - supports DOCX, DOC, XLSX, XLS, PDF"""
    if request.method == 'OPTIONS':
        return '', 204
    
    try:
        print("=" * 60, flush=True)
        print("Received Offer 2 template", flush=True)
        
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Validate file extension
        if not allowed_file(file.filename, ALLOWED_OFFER2_EXTENSIONS):
            return jsonify({
                'error': f'Unsupported template format. Allowed: {", ".join(ALLOWED_OFFER2_EXTENSIONS)}'
            }), 400
        
        file_extension = get_file_extension(file.filename)
        
        print(f"✓ Template format: {file_extension.upper()}", flush=True)
        
        # Save original file
        original_path = os.path.join(BASE_DIR, f'offer2_template_original.{file_extension}')
        file.save(original_path)
        
        print(f"✓ Template saved: {original_path}", flush=True)
        print(f"✓ Size: {os.path.getsize(original_path)} bytes", flush=True)
        
        # Convert to DOCX (required format for generation script)
        template_path = os.path.join(BASE_DIR, 'offer2_template.docx')
        
        if file_extension != 'docx':
            print(f"Converting {file_extension.upper()} template to DOCX...", flush=True)
            if not convert_to_docx(original_path, template_path, file_extension):
                return jsonify({
                    'error': f'Failed to convert {file_extension.upper()} to DOCX. LibreOffice may not be installed.',
                    'suggestion': 'Please upload a DOCX file instead or ensure LibreOffice is installed.'
                }), 500
        else:
            # Already DOCX, just copy
            shutil.copy(original_path, template_path)
        
        print(f"✓ Template ready for use: {template_path}", flush=True)
        print("=" * 60, flush=True)
        
        return jsonify({
            'success': True,
            'message': f'Template uploaded successfully ({file_extension.upper()})',
            'file_format': file_extension,
            'converted_to': 'docx'
        })
        
    except Exception as e:
        print(f"Error in upload-offer2: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/generate-offer', methods=['POST', 'OPTIONS'])
def api_generate_offer():
    if request.method == 'OPTIONS':
        return '', 204
    
    try:
        print("Starting offer generation...", flush=True)
        
        data = request.get_json() or {}
        markup = data.get('markup', 0)
        
        items_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        if not os.path.exists(items_path):
            return jsonify({'error': 'No items found. Please process Offer 1 first.'}), 400
        
        template_path = os.path.join(BASE_DIR, 'offer2_template.docx')
        if not os.path.exists(template_path):
            return jsonify({'error': 'No template found. Please upload Offer 2 template first.'}), 400
        
        if markup > 0:
            print(f"Applying {markup}% markup...", flush=True)
            with open(items_path, 'r', encoding='utf-8') as f:
                full_data = json.load(f)
            
            items = full_data.get('items', [])
            items = apply_markup_to_items(items, markup)
            full_data['items'] = items
            
            with open(items_path, 'w', encoding='utf-8') as f:
                json.dump(full_data, f, ensure_ascii=False, indent=2)
        
        print("Generating final offer...", flush=True)
        generate_script_path = os.path.join(BASE_DIR, 'generate_offer_doc.py')
        
        result = subprocess.run(
            ['python', generate_script_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=60
        )
        
        print(f"Generation returncode: {result.returncode}", flush=True)
        
        if result.returncode != 0:
            print(f"Generation failed: {result.stderr}", flush=True)
            return jsonify({
                'error': 'Offer generation failed',
                'details': result.stderr
            }), 500
        
        output_path = os.path.join(OUTPUT_FOLDER, 'final_offer1.docx')
        if not os.path.exists(output_path):
            return jsonify({'error': 'Output file not generated'}), 500
        
        with open(items_path, 'r', encoding='utf-8') as f:
            full_data = json.load(f)
            items = full_data.get('items', [])
        
        print("✓ Offer generated successfully", flush=True)
        
        return jsonify({
            'success': True,
            'message': 'Offer generated successfully',
            'download_url': '/api/download-offer',
            'items_count': len(items)
        })
        
    except Exception as e:
        print(f"Error in generate-offer: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/download-offer', methods=['GET', 'OPTIONS'])
def api_download_offer():
    if request.method == 'OPTIONS':
        return '', 204
    
    try:
        output_path = os.path.join(OUTPUT_FOLDER, 'final_offer1.docx')
        
        if not os.path.exists(output_path):
            return jsonify({'error': 'No offer generated yet'}), 404
        
        return send_file(
            output_path,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='requoted_offer.docx'
        )
        
    except Exception as e:
        print(f"Error in download-offer: {str(e)}", flush=True)
        return jsonify({'error': str(e)}), 500

def apply_markup_to_items(items, markup_percent):
    import re
    
    for item in items:
        price_str = str(item.get('unit_price', ''))
        if not price_str or price_str == '':
            price_str = str(item.get('price', ''))
        
        numbers = re.findall(r'\d+\.?\d*', price_str)
        if numbers:
            original_price = float(numbers[0])
            new_price = original_price * (1 + markup_percent / 100)
            currency = re.findall(r'[€$£¥]', price_str)
            currency_symbol = currency[0] if currency else '€'
            item['unit_price'] = currency_symbol + str(round(new_price, 2))
            if 'price' in item:
                item['price'] = currency_symbol + str(round(new_price, 2))
    
    return items

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print("=" * 60)
    print("Starting Requote AI Backend - SV9 Full MultiFormat")
    print(f"Server at: http://0.0.0.0:{port}")
    print(f"Supported Offer 1 formats: {', '.join(ALLOWED_OFFER1_EXTENSIONS)}")
    print(f"Supported Offer 2 formats: {', '.join(ALLOWED_OFFER2_EXTENSIONS)}")
    print("=" * 60)
    app.run(debug=True, host='0.0.0.0', port=port)