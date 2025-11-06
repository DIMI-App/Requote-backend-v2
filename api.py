from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import json
import subprocess
from werkzeug.utils import secure_filename
import threading
import time

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

# Thread-safe status storage
_status_lock = threading.Lock()
processing_status = {
    'status': 'idle',
    'message': '',
    'items_count': 0,
    'items': [],
    'started_at': None,
    'updated_at': None
}

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
        'version': 'SV7-Async-Fixed',
        'status': 'healthy'
    })

def process_pdf_background(filepath):
    global processing_status
    
    try:
        with _status_lock:
            processing_status['status'] = 'processing'
            processing_status['message'] = 'Starting PDF extraction...'
            processing_status['started_at'] = time.time()
            processing_status['updated_at'] = time.time()
        
        print("=== BACKGROUND THREAD STARTED ===", flush=True)
        print(f"Thread ID: {threading.current_thread().ident}", flush=True)
        print(f"Time: {time.strftime('%H:%M:%S')}", flush=True)
        
        items_output_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        extract_script_path = os.path.join(BASE_DIR, 'extract_pdf_direct.py')
        
        print(f"Script path: {extract_script_path}", flush=True)
        print(f"Script exists: {os.path.exists(extract_script_path)}", flush=True)
        
        with _status_lock:
            processing_status['message'] = 'Converting PDF pages to images...'
            processing_status['updated_at'] = time.time()
        
        print("Starting subprocess...", flush=True)
        
        result = subprocess.run(
            ['python', extract_script_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=300
        )
        
        print(f"Subprocess completed with code: {result.returncode}", flush=True)
        print(f"STDOUT length: {len(result.stdout)} chars", flush=True)
        print(f"STDERR length: {len(result.stderr)} chars", flush=True)
        
        if result.stdout:
            print("=== SUBPROCESS STDOUT ===", flush=True)
            print(result.stdout, flush=True)
        
        if result.stderr:
            print("=== SUBPROCESS STDERR ===", flush=True)
            print(result.stderr, flush=True)
        
        if result.returncode != 0:
            with _status_lock:
                processing_status['status'] = 'error'
                processing_status['message'] = 'Extraction failed: ' + (result.stderr or result.stdout or 'Unknown error')
                processing_status['updated_at'] = time.time()
            print("❌ Subprocess failed", flush=True)
            return
        
        print("Checking for output file...", flush=True)
        
        if not os.path.exists(items_output_path):
            with _status_lock:
                processing_status['status'] = 'error'
                processing_status['message'] = 'Items file not created after extraction'
                processing_status['updated_at'] = time.time()
            print(f"❌ Output file not found: {items_output_path}", flush=True)
            return
        
        print(f"✓ Output file exists: {items_output_path}", flush=True)
        
        with _status_lock:
            processing_status['message'] = 'Loading extracted items...'
            processing_status['updated_at'] = time.time()
        
        with open(items_output_path, 'r', encoding='utf-8') as f:
            full_data = json.load(f)
        
        items = full_data.get('items', [])
        
        print(f"✓ Loaded {len(items)} items from file", flush=True)
        
        with _status_lock:
            processing_status['status'] = 'completed'
            processing_status['message'] = f'Successfully extracted {len(items)} items'
            processing_status['items_count'] = len(items)
            processing_status['items'] = items
            processing_status['updated_at'] = time.time()
        
        print("=== BACKGROUND THREAD COMPLETED ===", flush=True)
        elapsed = time.time() - processing_status['started_at']
        print(f"Total time: {elapsed:.1f} seconds", flush=True)
        
    except Exception as e:
        print(f"=== BACKGROUND THREAD ERROR: {str(e)} ===", flush=True)
        import traceback
        traceback.print_exc()
        
        with _status_lock:
            processing_status['status'] = 'error'
            processing_status['message'] = str(e)
            processing_status['updated_at'] = time.time()

@app.route('/api/process-offer1', methods=['POST', 'OPTIONS'])
def api_process_offer1():
    global processing_status
    
    if request.method == 'OPTIONS':
        return '', 204
    
    try:
        print("=" * 60, flush=True)
        print("Received request to process Offer 1", flush=True)
        print(f"Time: {time.strftime('%H:%M:%S')}", flush=True)
        
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, 'offer1.pdf')
        file.save(filepath)
        
        print(f"✓ File saved: {filepath}", flush=True)
        print(f"✓ File size: {os.path.getsize(filepath)} bytes", flush=True)
        
        # Reset status
        with _status_lock:
            processing_status = {
                'status': 'processing',
                'message': 'File uploaded, starting extraction...',
                'items_count': 0,
                'items': [],
                'started_at': time.time(),
                'updated_at': time.time()
            }
        
        print("Starting background thread...", flush=True)
        
        # Start background thread
        thread = threading.Thread(target=process_pdf_background, args=(filepath,), daemon=True)
        thread.start()
        
        print(f"✓ Background thread started: {thread.ident}", flush=True)
        print("=" * 60, flush=True)
        
        # Return immediately
        return jsonify({
            'success': True,
            'message': 'Processing started. Poll /api/status for updates.',
            'status': 'processing'
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
    if request.method == 'OPTIONS':
        return '', 204
    
    try:
        print("Received Offer 2 template", flush=True)
        
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        filepath = os.path.join(BASE_DIR, 'offer2_template.docx')
        file.save(filepath)
        
        print(f"✓ Template saved: {filepath}", flush=True)
        
        return jsonify({
            'success': True,
            'message': 'Template uploaded successfully'
        })
        
    except Exception as e:
        print(f"Error in upload-offer2: {str(e)}", flush=True)
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
    print("Starting Requote AI Backend - SV7 Async Fixed")
    print(f"Server at: http://0.0.0.0:{port}")
    print("=" * 60)
    app.run(debug=True, host='0.0.0.0', port=port)