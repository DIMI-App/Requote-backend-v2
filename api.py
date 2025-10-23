from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import json
import subprocess
from werkzeug.utils import secure_filename

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
        'version': 'Day13-Clean',
        'status': 'healthy'
    })

@app.route('/api/process-offer1', methods=['POST', 'OPTIONS'])
def api_process_offer1():
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
        
        print("STEP 1: Processing with Document AI...")
        test_process_path = os.path.join(BASE_DIR, 'test_process.py')
        result = subprocess.run(
            ['python', test_process_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=180
        )
        
        if result.returncode != 0:
            print("Document AI Error")
            return jsonify({
                'error': 'Document AI processing failed',
                'details': result.stderr
            }), 500
        
        print("Document AI complete")
        
        extracted_text_path = os.path.join(OUTPUT_FOLDER, 'extracted_text.txt')
        
        if not os.path.exists(extracted_text_path):
            print("Extracted text file not found")
            return jsonify({'error': 'Extracted text file not found'}), 500
        
        print("STEP 2: Extracting items with OpenAI...")
        
        items_output_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        extract_items_path = os.path.join(BASE_DIR, 'extract_items.py')
        
        result = subprocess.run(
            ['python', extract_items_path, extracted_text_path, items_output_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=120
        )
        
        if result.returncode != 0:
            print("Item extraction failed")
            return jsonify({
                'error': 'Item extraction failed',
                'details': result.stderr
            }), 500
        
        print("Extraction complete")
        
        if not os.path.exists(items_output_path):
            return jsonify({'error': 'Items file not created'}), 500
        
        with open(items_output_path, 'r', encoding='utf-8') as f:
            full_data = json.load(f)
        
        items = full_data.get('items', [])
        
        print("SUCCESS! Extracted " + str(len(items)) + " items")
        
        return jsonify({
            'success': True,
            'items_count': len(items),
            'items': items,
            'message': 'Successfully extracted ' + str(len(items)) + ' items'
        })
        
    except Exception as e:
        print("CRITICAL ERROR: " + str(e))
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

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
        
        try:
            result = subprocess.run(
                ['python', generate_script_path],
                capture_output=True,
                text=True,
                cwd=BASE_DIR,
                timeout=60
            )
            
            print("=" * 60)
            print("GENERATE OFFER OUTPUT:")
            print("Return code: " + str(result.returncode))
            print("STDOUT:")
            print(result.stdout)
            print("STDERR:")
            print(result.stderr)
            print("=" * 60)
            
            if result.returncode != 0:
                error_msg = result.stderr if result.stderr else result.stdout
                print("Offer generation failed with return code: " + str(result.returncode))
                print("Full error: " + error_msg)
                return jsonify({
                    'error': 'Offer generation failed',
                    'return_code': result.returncode,
                    'stderr': result.stderr,
                    'stdout': result.stdout,
                    'full_output': error_msg
                }), 500
        except subprocess.TimeoutExpired:
            print("Offer generation timed out after 60 seconds")
            return jsonify({'error': 'Offer generation timed out'}), 500
        except Exception as e:
            print("Exception during offer generation: " + str(e))
            import traceback
            traceback.print_exc()
            return jsonify({'error': 'Exception: ' + str(e)}), 500
        
        output_path = os.path.join(OUTPUT_FOLDER, 'final_offer1.docx')
        if not os.path.exists(output_path):
            return jsonify({'error': 'Output file not generated'}), 500
        
        with open(items_path, 'r', encoding='utf-8') as f:
            full_data = json.load(f)
            items = full_data.get('items', [])
        
        print("Final offer generated successfully")
        
        return jsonify({
            'success': True,
            'message': 'Offer generated successfully',
            'download_url': '/api/download-offer',
            'items_count': len(items)
        })
        
    except Exception as e:
        print("Error: " + str(e))
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
    print("Starting Requote AI Backend Server")
    print("Server will be available at: http://0.0.0.0:" + str(port))
    app.run(debug=True, host='0.0.0.0', port=port)