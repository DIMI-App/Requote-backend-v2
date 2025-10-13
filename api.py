from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import json
import subprocess
from werkzeug.utils import secure_filename

app = Flask(__name__)

# CORS configuration
CORS(app, resources={
    r"/*": {
        "origins": "*",
        "methods": ["GET", "POST", "OPTIONS"],
        "allow_headers": ["Content-Type"]
    }
})

# Get the absolute path of the project directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

# ============================================
# ENDPOINT 1: Health Check
# ============================================
@app.route('/', methods=['GET'])
def home():
    return jsonify({
        'message': 'Requote AI Backend is running!',
        'version': 'SV3',
        'status': 'healthy'
    })

# ============================================
# ENDPOINT 2: Upload and Process Offer 1
# ============================================
@app.route('/api/process-offer1', methods=['POST'])
def api_process_offer1():
    try:
        print("📤 Received request to process Offer 1")
        
        # Check if file was uploaded
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Save uploaded file with absolute path
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, 'offer1.pdf')
        file.save(filepath)
        
        print(f"✅ File saved: {filepath}")
        
        # Step 1: Process with Document AI
        print("🔍 Processing with Document AI...")
        test_process_path = os.path.join(BASE_DIR, 'test_process.py')
        result = subprocess.run(
            ['python', test_process_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR
        )
        
        if result.returncode != 0:
            print(f"❌ Document AI Error: {result.stderr}")
            return jsonify({'error': 'Document AI processing failed', 'details': result.stderr}), 500
        
        print("✅ Document AI complete")
        
        # Step 2: Extract items with OpenAI
        print("🤖 Extracting items with OpenAI...")
        
        # Define absolute paths for input and output
        extracted_text_path = os.path.join(OUTPUT_FOLDER, 'extracted_text.txt')
        items_output_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        extract_items_path = os.path.join(BASE_DIR, 'extract_items.py')
        
        # Verify extracted text file exists
        if not os.path.exists(extracted_text_path):
            print(f"❌ Extracted text file not found: {extracted_text_path}")
            return jsonify({'error': 'Extracted text file not found'}), 500
        
        print(f"✅ Found extracted text file")
        
        result = subprocess.run(
            ['python', extract_items_path, extracted_text_path, items_output_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR
        )
        
        print(f"STDOUT: {result.stdout}")
        if result.stderr:
            print(f"STDERR: {result.stderr}")
        
        if result.returncode != 0:
            print(f"❌ OpenAI extraction failed")
            return jsonify({'error': 'Item extraction failed', 'details': result.stderr}), 500
        
        print("✅ Extraction complete")
        
        # Step 3: Load extracted items
        if not os.path.exists(items_output_path):
            print(f"❌ Items file not found at: {items_output_path}")
            return jsonify({'error': 'Items file not created'}), 500
        
        with open(items_output_path, 'r', encoding='utf-8') as f:
            full_data = json.load(f)
        
        items = full_data.get('items', [])
        
        print(f"✅ Extracted {len(items)} items")
        
        return jsonify({
            'success': True,
            'items_count': len(items),
            'items': items,
            'full_data': full_data,
            'message': f'Successfully extracted {len(items)} items'
        })
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

# ============================================
# ENDPOINT 3: Upload Offer 2 (Template)
# ============================================
@app.route('/api/upload-offer2', methods=['POST'])
def api_upload_offer2():
    try:
        print("📤 Received Offer 2 template")
        
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Save as offer2_template.docx
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

# ============================================
# ENDPOINT 4: Generate Final Offer
# ============================================
@app.route('/api/generate-offer', methods=['POST'])
def api_generate_offer():
    try:
        print("🔄 Starting offer generation...")
        
        # Get optional parameters
        data = request.get_json() or {}
        markup = data.get('markup', 0)
        
        # Check if items exist
        items_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        if not os.path.exists(items_path):
            return jsonify({'error': 'No items found. Please process Offer 1 first.'}), 400
        
        # Apply markup if requested
        if markup > 0:
            print(f"💰 Applying {markup}% markup...")
            with open(items_path, 'r', encoding='utf-8') as f:
                full_data = json.load(f)
            
            items = full_data.get('items', [])
            items = apply_markup_to_items(items, markup)
            full_data['items'] = items
            
            # Save updated items
            with open(items_path, 'w', encoding='utf-8') as f:
                json.dump(full_data, f, ensure_ascii=False, indent=2)
        
        # Run the generation script
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
        
        # Check if output file exists
        output_path = os.path.join(OUTPUT_FOLDER, 'final_offer1.docx')
        if not os.path.exists(output_path):
            return jsonify({'error': 'Output file not generated'}), 500
        
        # Load items count for response
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

# ============================================
# ENDPOINT 5: Download Generated Offer
# ============================================
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

# ============================================
# HELPER FUNCTION: Apply Markup
# ============================================
def apply_markup_to_items(items, markup_percent):
    """Apply percentage markup to all item prices"""
    import re
    
    for item in items:
        price_str = str(item.get('unit_price', ''))
        if not price_str or price_str == '':
            price_str = str(item.get('price', ''))
        
        # Extract numeric value
        numbers = re.findall(r'\d+\.?\d*', price_str)
        if numbers:
            original_price = float(numbers[0])
            new_price = original_price * (1 + markup_percent / 100)
            # Replace price in original format
            currency = re.findall(r'[€$£¥]', price_str)
            currency_symbol = currency[0] if currency else '€'
            item['unit_price'] = f"{currency_symbol}{new_price:.2f}"
            if 'price' in item:
                item['price'] = f"{currency_symbol}{new_price:.2f}"
    
    return items

# ============================================
# RUN SERVER
# ============================================
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print("🚀 Starting Requote AI Backend Server (SV3)...")
    print(f"📡 Server will be available at: http://0.0.0.0:{port}")
    app.run(debug=True, host='0.0.0.0', port=port)