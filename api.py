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
        "methods": ["GET", "POST", "OPTIONS"],
        "allow_headers": ["Content-Type"]
    }
})

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

@app.route('/', methods=['GET'])
def home():
    return jsonify({
        'message': 'Requote AI Backend is running!',
        'version': 'SV3.1-Day13',
        'status': 'healthy'
    })

@app.route('/api/process-offer1', methods=['POST'])
def api_process_offer1():
    try:
        print("=" * 70)
        print("📤 Received request to process Offer 1")
        print("=" * 70)
        
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, 'offer1.pdf')
        file.save(filepath)
        
        print(f"✅ File saved: {filepath}")
        print(f"📊 File size: {os.path.getsize(filepath)} bytes")
        
        # === STEP 1: Document AI Processing ===
        print("\n" + "=" * 70)
        print("🔍 STEP 1: Processing with Document AI...")
        print("=" * 70)
        
        test_process_path = os.path.join(BASE_DIR, 'test_process.py')
        result = subprocess.run(
            ['python', test_process_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR
        )
        
        print("STDOUT:", result.stdout)
        if result.stderr:
            print("STDERR:", result.stderr)
        
        if result.returncode != 0:
            print(f"❌ Document AI Error (exit code {result.returncode})")
            return jsonify({
                'error': 'Document AI processing failed',
                'details': result.stderr,
                'stdout': result.stdout
            }), 500
        
        print("✅ Document AI complete")
        
        # === STEP 2: Verify extracted_text.txt exists ===
        extracted_text_path = os.path.join(OUTPUT_FOLDER, 'extracted_text.txt')
        
        print("\n" + "=" * 70)
        print("🔍 STEP 2: Verifying extracted text file...")
        print("=" * 70)
        print(f"📂 Looking for: {extracted_text_path}")
        print(f"📂 File exists: {os.path.exists(extracted_text_path)}")
        
        if os.path.exists(extracted_text_path):
            file_size = os.path.getsize(extracted_text_path)
            print(f"✅ File found! Size: {file_size} bytes")
            
            # Preview content
            with open(extracted_text_path, 'r', encoding='utf-8') as f:
                preview = f.read(200)
                print(f"📄 Preview: {preview[:100]}...")
        else:
            print("❌ FILE NOT FOUND!")
            print(f"📁 Output folder contents: {os.listdir(OUTPUT_FOLDER)}")
            return jsonify({
                'error': 'Extracted text file not found',
                'expected_path': extracted_text_path,
                'output_folder_contents': os.listdir(OUTPUT_FOLDER)
            }), 500
        
        # === STEP 3: OpenAI Item Extraction ===
        print("\n" + "=" * 70)
        print("🤖 STEP 3: Extracting items with OpenAI...")
        print("=" * 70)
        
        items_output_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        extract_items_path = os.path.join(BASE_DIR, 'extract_items.py')
        
        result = subprocess.run(
            ['python', extract_items_path, extracted_text_path, items_output_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR
        )
        
        print("STDOUT:", result.stdout)
        if result.stderr:
            print("STDERR:", result.stderr)
        
        if result.returncode != 0:
            print(f"❌ Item extraction failed (exit code {result.returncode})")
            return jsonify({
                'error': 'Item extraction failed',
                'details': result.stderr,
                'stdout': result.stdout
            }), 500
        
        print("✅ Extraction complete")
        
        # === STEP 4: Load and return results ===
        if not os.path.exists(items_output_path):
            return jsonify({'error': 'Items file not created'}), 500
        
        with open(items_output_path, 'r', encoding='utf-8') as f:
            full_data = json.load(f)
        
        items = full_data.get('items', [])
        
        print("\n" + "=" * 70)
        print(f"✅ SUCCESS! Extracted {len(items)} items")
        print("=" * 70)
        
        return jsonify({
            'success': True,
            'items_count': len(items),
            'items': items,
            'full_data': full_data,
            'message': f'Successfully extracted {len(items)} items'
        })
        
    except Exception as e:
        print(f"\n❌ CRITICAL ERROR: {str(e)}")
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
        print("🔄 Starting offer generation...")
        
        data = request.get_json() or {}
        markup = data.get('markup', 0)
        
        items_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        if not os.path.exists(items_path):
            return jsonify({'error': 'No items found. Please process Offer 1 first.'}), 400
        
        if markup > 0:
            print(f"💰 Applying {markup}% markup...")
            with open(items_path, 'r', encoding='utf-8') as f:
                full_data = json.load(f)
            
            items = full_data.get('items', [])
            items = apply_markup_to_items(items, markup)
            full_data['items'] = items
            
            with open(items_path, 'w', encoding='utf-8') as f:
                json.dump(full_data, f, ensure_ascii=False, indent=2)
        
        print("📝 Generating final offer...")
        generate_script_path = os.path.join(BASE_DIR, 'generate_offer_doc.py')
        result = subprocess.run(
            ['python', generate_script_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR
        )
        
        if result.returncode != 0:
            print(f"Error: {result.stderr}")
            return jsonify({'error': 'Offer generation failed', 'details': result.stderr}), 500
        
        print(result.stdout)
        
        output_path = os.path.join(OUTPUT_FOLDER, 'final_offer1.docx')
        if not os.path.exists(output_path):
            return jsonify({'error': 'Output file not generated'}), 500
        
        with open(items_path, 'r', encoding='utf-8') as f:
            full_data = json.load(f)
            items = full_data.get('items', [])
        
        print(f"✅ Final offer generated successfully")
        
        return jsonify({
            'success': True,
            'message': 'Offer generated successfully',
            'download_url': '/api/download-offer',
            'items_count': len(items)
        })
        
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
        if not price_str or price_str == '':
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
    print("=" * 70)
    print("🚀 Starting Requote AI Backend Server (Day 13 - v3.1)")
    print("=" * 70)
    print(f"📡 Server will be available at: http://0.0.0.0:{port}")
    print("=" * 70)
    app.run(debug=True, host='0.0.0.0', port=port)