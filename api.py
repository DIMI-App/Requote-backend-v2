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

# Store processing status
processing_status = {
    'status': 'idle',
    'message': '',
    'items_count': 0,
    'items': []
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
        'version': 'SV7-Async',
        'status': 'healthy'
    })

def process_pdf_background(filepath):
    global processing_status
    
    try:
        processing_status['status'] = 'processing'
        processing_status['message'] = 'Extracting items from PDF...'
        
        print("=== BACKGROUND THREAD STARTED ===", flush=True)
        
        items_output_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        extract_script_path = os.path.join(BASE_DIR, 'extract_pdf_direct.py')
        
        print(f"Script path: {extract_script_path}", flush=True)
        print(f"Script exists: {os.path.exists(extract_script_path)}", flush=True)
        
        print("Starting subprocess...", flush=True)
        
        result = subprocess.run(
            ['python', extract_script_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=300
        )
        
        print(f"Subprocess completed with code: {result.returncode}", flush=True)
        print(f"STDOUT: {result.stdout}", flush=True)
        print(f"STDERR: {result.stderr}", flush=True)
        
        if result.returncode != 0:
            processing_status['status'] = 'error'
            processing_status['message'] = 'Extraction failed: ' + result.stderr
            return
        
        if not os.path.exists(items_output_path):
            processing_status['status'] = 'error'
            processing_status['message'] = 'Items file not created'
            return
        
        with open(items_output_path, 'r', encoding='utf-8') as f:
            full_data = json.load(f)
        
        items = full_data.get('items', [])
        
        processing_status['status'] = 'completed'
        processing_status['message'] = f'Successfully extracted {len(items)} items'
        processing_status['items_count'] = len(items)
        processing_status['items'] = items
        
        print("=== BACKGROUND THREAD COMPLETED ===", flush=True)
        
    except Exception as e:
        print(f"=== BACKGROUND THREAD ERROR: {str(e)} ===", flush=True)
        import traceback
        traceback.print_exc()
        processing_status['status'] = 'error'
        processing_status['message'] = str(e)

@app.route('/api/process-offer1', methods=['POST', 'OPTIONS'])
def api_process_offer1():
    global processing_status
    
    if request.method == 'OPTIONS':
        return '', 204
    
    try:
        print("Received request to process Offer 1")
        
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, 'offer1.pdf')
        file.save(filepath)
        
        print("File saved: " + filepath)
        
        # Reset status
        processing_status = {
            'status': 'processing',
            'message': 'Starting extraction...',
            'items_count': 0,
            'items': []
        }
        
        # Start background thread
        thread = threading.Thread(target=process_pdf_background, args=(filepath,))
        thread.daemon = True
        thread.start()
        
        # Return immediately
        return jsonify({
            'success': True,
            'message': 'Processing started. Poll /api/status for updates.',
            'status': 'processing'
        })
        
    except Exception as e:
        print("ERROR: " + str(e))
        return jsonify({'error': str(e)}), 500

@app.route('/api/status', methods=['GET', 'OPTIONS'])
def api_status():
    if request.method == 'OPTIONS':
        return '', 204
    
    return jsonify(processing_status)

@app.route('/api/upload-offer2', methods=['POST', 'OPTIONS'])
def api_upload_offer2():
    if request.method == 'OPTIONS':
        return '', 204
    
    try:
        print("Received Offer 2 template")
        
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        filepath = os.path.join(BASE_DIR, 'offer2_template.docx')
        file.save(filepath)
        
        print("Template saved")
        
        return jsonify({
            'success': True,
            'message': 'Template uploaded successfully'
        })
        
    except Exception as e:
        print("Error: " + str(e))
        return jsonify({'error': str(e)}), 500

@app.route('/api/generate-offer', methods=['POST', 'OPTIONS'])
def api_generate_offer():
    if request.method == 'OPTIONS':
        return '', 204
    
    try:
        print("Starting offer generation...")
        
        data = request.get_json() or {}
        markup = data.get('markup', 0)
        
        items_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        if not os.path.exists(items_path):
            return jsonify({'error': 'No items found. Please process Offer 1 first.'}), 400
        
        if markup > 0:
            print("Applying markup...")
            with open(items_path, 'r', encoding='utf-8') as f:
                full_data = json.load(f)
            
            items = full_data.get('items', [])
            items = apply_markup_to_items(items, markup)
            full_data['items'] = items
            
            with open(items_path, 'w', encoding='utf-8') as f:
                json.dump(full_data, f, ensure_ascii=False, indent=2)
        
        print("Generating final offer...")
        generate_script_path = os.path.join(BASE_DIR, 'generate_offer_doc.py')
        
        result = subprocess.run(
            ['python', generate_script_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=60
        )
        
        if result.returncode != 0:
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
        
        return jsonify({
            'success': True,
            'message': 'Offer generated successfully',
            'download_url': '/api/download-offer',
            'items_count': len(items)
        })
        
    except Exception as e:
        print("Error: " + str(e))
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
        print("Error: " + str(e))
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
    print("Starting Requote AI Backend - SV7 Async")
    print("Server at: http://0.0.0.0:" + str(port))
    app.run(debug=True, host='0.0.0.0', port=port)