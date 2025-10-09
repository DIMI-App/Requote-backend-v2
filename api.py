from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import json
import subprocess
from werkzeug.utils import secure_filename
from docx import Document

app = Flask(__name__)
CORS(app)  # Allow requests from Lovable

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
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
        'version': '1.0.0',
        'status': 'healthy'
    })

# ============================================
# ENDPOINT 2: Upload and Process Offer 1
# ============================================
@app.route('/api/process-offer1', methods=['POST'])
def api_process_offer1():
    """
    Upload Offer 1 (supplier PDF)
    Returns extracted items as JSON
    """
    try:
        print("📤 Received request to process Offer 1")
        
        # Check if file was uploaded
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], 'offer1.pdf')
        file.save(filepath)
        
        print(f"✅ File saved: {filepath}")
        
        # Step 1: Process with Document AI (call your existing script)
        print("🔍 Processing with Document AI...")
        result = subprocess.run(
            ['python', 'test_process.py'],
            capture_output=True,
            text=True
        )
        
        if result.returncode != 0:
            print(f"Error: {result.stderr}")
            return jsonify({'error': 'Document AI processing failed'}), 500
        
        print("✅ Document AI complete")
        
        # Step 2: Extract items with OpenAI (call your existing script)
        print("🤖 Extracting items with OpenAI...")
        result = subprocess.run(
            ['python', 'extract_items.py'],
            capture_output=True,
            text=True
        )
        
        if result.returncode != 0:
            print(f"Error: {result.stderr}")
            return jsonify({'error': 'Item extraction failed'}), 500
        
        print("✅ Extraction complete")
        
        # Load extracted items
        items_path = "outputs/items_offer1.json"
        with open(items_path, 'r', encoding='utf-8') as f:
            items = json.load(f)
        
        print(f"✅ Extracted {len(items)} items")
        
        return jsonify({
            'success': True,
            'items_count': len(items),
            'items': items,
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
    """
    Upload Offer 2 (company template)
    Just saves it for later use
    """
    try:
        print("📤 Received Offer 2 template")
        
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Save as offer2_template.docx
        filepath = "offer2_template.docx"
        file.save(filepath)
        
        print(f"✅ Template saved: {filepath}")
        
        return jsonify({
            'success': True,
            'message': 'Template uploaded successfully',
            'filename': filepath
        })
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return jsonify({'error': str(e)}), 500

# ============================================
# ENDPOINT 4: Generate Final Offer
# ============================================
@app.route('/api/generate-offer', methods=['POST'])
def api_generate_offer():
    """
    Merge Offer 1 items + Offer 2 template
    Returns downloadable DOCX
    """
    try:
        print("🔄 Starting offer generation...")
        
        # Get optional parameters
        data = request.get_json() or {}
        markup = data.get('markup', 0)  # Percentage markup
        
        # Check if items exist
        items_path = "outputs/items_offer1.json"
        if not os.path.exists(items_path):
            return jsonify({'error': 'No items found. Please process Offer 1 first.'}), 400
        
        # Apply markup if requested
        if markup > 0:
            print(f"💰 Applying {markup}% markup...")
            with open(items_path, 'r', encoding='utf-8') as f:
                items = json.load(f)
            
            items = apply_markup_to_items(items, markup)
            
            # Save updated items
            with open(items_path, 'w', encoding='utf-8') as f:
                json.dump(items, f, ensure_ascii=False, indent=2)
        
        # Run the generation script
        print("📝 Generating final offer...")
        result = subprocess.run(
            ['python', 'generate_offer_doc.py'],
            capture_output=True,
            text=True
        )
        
        if result.returncode != 0:
            print(f"Error: {result.stderr}")
            return jsonify({'error': 'Offer generation failed'}), 500
        
        print(result.stdout)  # Show the success message
        
        # Check if output file exists
        output_path = "outputs/final_offer1.docx"
        if not os.path.exists(output_path):
            return jsonify({'error': 'Output file not generated'}), 500
        
        # Load items count for response
        with open(items_path, 'r', encoding='utf-8') as f:
            items = json.load(f)
        
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
    """
    Download the generated final offer
    """
    try:
        output_path = "outputs/final_offer1.docx"
        
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
        price_str = str(item.get('price', ''))
        # Extract numeric value
        numbers = re.findall(r'\d+\.?\d*', price_str)
        if numbers:
            original_price = float(numbers[0])
            new_price = original_price * (1 + markup_percent / 100)
            # Replace price in original format
            currency = re.findall(r'[€$£¥]', price_str)
            currency_symbol = currency[0] if currency else '€'
            item['price'] = f"{currency_symbol}{new_price:.2f}"
    
    return items

# ============================================
# RUN SERVER
# ============================================
if __name__ == '__main__':
    print("🚀 Starting Requote AI Backend Server...")
    print("📡 Server will be available at: http://localhost:5000")
    print("🌐 Ready to receive requests from Lovable frontend!\n")
    app.run(debug=True, host='0.0.0.0', port=5000)