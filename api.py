from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import json
import subprocess
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge

from extract_items import extract_items_from_text
from process_offer1 import extract_offer1_text, save_text_to_file

from extract_items import extract_items_from_text
from process_offer1 import extract_offer1_text, save_text_to_file

app = Flask(__name__)

CORS(app, resources={
    r"/*": {
        "origins": "*",
        "methods": ["GET", "POST", "OPTIONS"],
        "allow_headers": [
            "Content-Type",
            "access-control-allow-origin",
            "Access-Control-Allow-Origin",
        ],
    }
})

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024


@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(error):
    return jsonify({'error': 'Uploaded file is too large'}), 413

@app.route('/', methods=['GET'])
def home():
    return jsonify({
        'message': 'Requote AI Backend is running!',
        'version': '1.0.0',
        'status': 'healthy'
    })

@app.route('/api/process-offer1', methods=['POST'])
def api_process_offer1():
    try:
        print("📤 Received request to process Offer 1")
        
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, 'offer1.pdf')
        file.save(filepath)
        
        print(f"✅ File saved: {filepath}")
        
        print("🔍 Processing Offer 1 with Document AI (with fallback)...")
        extracted_text, diagnostics = extract_offer1_text(filepath)

        if not extracted_text.strip():
            error_payload = {'error': 'Failed to extract text from Offer 1'}
            if diagnostics:
                error_payload['details'] = diagnostics
                doc_error = diagnostics.get('document_ai_error')
                if doc_error and doc_error.get('type') in {
                    'document_ai_permission',
                    'document_ai_unauthenticated',
                    'document_ai_credentials',
                }:
                    return jsonify(error_payload), 503
            return jsonify(error_payload), 500
        extracted_text = extract_offer1_text(filepath)

        if not extracted_text.strip():
            return jsonify({'error': 'Failed to extract text from Offer 1'}), 500

        extracted_text_path = os.path.join(OUTPUT_FOLDER, 'extracted_text.txt')
        save_text_to_file(extracted_text, extracted_text_path)

        print("✅ Text extraction complete")

        print("🤖 Extracting items with OpenAI...")

        items_output_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        success, error_info = extract_items_from_text(extracted_text, items_output_path)

        if not success:
            error_response = {
                'error': 'Item extraction failed',
            }

            if error_info:
                error_response['details'] = error_info

                if error_info.get('type') == 'openai_error' and error_info.get('status') == 429:
                    error_response['error'] = 'OpenAI quota exceeded'
                    return jsonify(error_response), 429

            return jsonify(error_response), 500
        success = extract_items_from_text(extracted_text, items_output_path)

        if not success:
            return jsonify({'error': 'Item extraction failed'}), 500

        print("✅ Extraction complete")
        
        if not os.path.exists(items_output_path):
            return jsonify({'error': 'Items file not created'}), 500
        
        with open(items_output_path, 'r', encoding='utf-8') as f:
            full_data = json.load(f)
        
        items = full_data.get('items', [])
        
        print(f"✅ Extracted {len(items)} items")
        
        response_payload = {
            'success': True,
            'items_count': len(items),
            'items': items,
            'full_data': full_data,
            'message': f'Successfully extracted {len(items)} items'
        }

        if diagnostics:
            response_payload['diagnostics'] = diagnostics

        return jsonify(response_payload)
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/upload-offer2', methods=['POST'])
def api_upload_offer2():
    try:
        print("📤 Received Offer 2 template")
        
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        filepath = os.path.join(BASE_DIR, 'offer2_template.docx')
        file.save(filepath)
        
        print(f"✅ Template saved: {filepath}")
        
        return jsonify({
            'success': True,
            'message': 'Template uploaded successfully',
            'filename': 'offer2_template.docx'
        })
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/generate-offer', methods=['POST'])
def api_generate_offer():
    try:
        print("=" * 60)
        print("🔄 GENERATE OFFER - DETAILED DEBUG")
        print("=" * 60)
        
        data = request.get_json() or {}
        markup = data.get('markup', 0)
        
        items_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        template_path = os.path.join(BASE_DIR, 'offer2_template.docx')
        
        print(f"\n📋 Checking files...")
        print(f"   Items: {os.path.exists(items_path)}")
        print(f"   Template: {os.path.exists(template_path)}")
        
        if not os.path.exists(items_path):
            return jsonify({'error': 'No items found. Process Offer 1 first.'}), 400
        
        if not os.path.exists(template_path):
            return jsonify({'error': 'No template found. Upload Offer 2 first.'}), 400
        
        with open(items_path, 'r', encoding='utf-8') as f:
            full_data = json.load(f)
        
        items = full_data.get('items', [])
        print(f"   Items count: {len(items)}")
        
        if len(items) == 0:
            return jsonify({'error': 'Items array is empty'}), 400
        
        if markup > 0:
            print(f"\n💰 Applying {markup}% markup...")
            items = apply_markup_to_items(items, markup)
            full_data['items'] = items
            
            with open(items_path, 'w', encoding='utf-8') as f:
                json.dump(full_data, f, ensure_ascii=False, indent=2)
        
        print("\n📝 Running generation script...")
        generate_script_path = os.path.join(BASE_DIR, 'generate_offer_doc.py')
        
        result = subprocess.run(
            ['python', generate_script_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=30
        )
        
        print(f"   Return code: {result.returncode}")
        
        if result.stdout:
            print(f"\n📤 STDOUT:\n{result.stdout}")
        
        if result.stderr:
            print(f"\n📤 STDERR:\n{result.stderr}")
        
        if result.returncode != 0:
            error_msg = result.stderr or result.stdout or "Unknown error"
            return jsonify({
                'error': 'Generation failed',
                'details': error_msg
            }), 500
        
        output_path = os.path.join(OUTPUT_FOLDER, 'final_offer1.docx')
        
        if not os.path.exists(output_path):
            return jsonify({'error': 'Output file not created'}), 500
        
        print(f"\n✅ SUCCESS")
        
        return jsonify({
            'success': True,
            'message': 'Offer generated successfully',
            'download_url': '/api/download-offer',
            'items_count': len(items)
        })
        
    except subprocess.TimeoutExpired:
        return jsonify({'error': 'Generation timed out'}), 500
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/download-offer', methods=['GET'])
def api_download_offer():
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
        print(f"❌ Error: {str(e)}")
        return jsonify({'error': str(e)}), 500

def apply_markup_to_items(items, markup_percent):
    import re
    
    for item in items:
        price_str = str(item.get('unit_price', ''))
        if not price_str:
            price_str = str(item.get('price', ''))
        
        numbers = re.findall(r'\d+\.?\d*', price_str)
        if numbers:
            original_price = float(numbers[0])
            new_price = original_price * (1 + markup_percent / 100)
            currency = re.findall(r'[€$£¥]', price_str)
            currency_symbol = currency[0] if currency else '€'
            item['unit_price'] = f"{currency_symbol}{new_price:.2f}"
            if 'price' in item:
                item['price'] = f"{currency_symbol}{new_price:.2f}"
    
    return items

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print("🚀 Starting Requote AI Backend...")
    app.run(debug=True, host='0.0.0.0', port=port)