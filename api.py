from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import json
import subprocess
from werkzeug.utils import secure_filename
from docx import Document

app = Flask(__name__)
CORS(app, resources={
    r"/*": {
        "origins": "*",
        "methods": ["GET", "POST", "OPTIONS"],
        "allow_headers": ["Content-Type"]
    }
})  # Allow requests from Lovable

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
        print("=" * 60)
        print("📤 NEW REQUEST: PROCESS OFFER 1")
        print("=" * 60)
        
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # === FORCE OVERWRITE OLD PDF ===
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, 'offer1.pdf')
        
        print(f"💾 Saving new PDF: {filepath}")
        file.save(filepath)  # This overwrites automatically
        
        file_size = os.path.getsize(filepath)
        print(f"✅ File saved: {file_size} bytes")
        
        # === FORCE DELETE OLD EXTRACTION FILES ===
        print("\n🧹 Force clearing old extraction data...")
        
        extracted_text_path = os.path.join(OUTPUT_FOLDER, 'extracted_text.txt')
        items_output_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        
        # Delete old files if they exist
        for old_file in [extracted_text_path, items_output_path]:
            if os.path.exists(old_file):
                try:
                    os.remove(old_file)
                    print(f"   ✓ Deleted: {os.path.basename(old_file)}")
                except Exception as e:
                    print(f"   ⚠️  Could not delete {os.path.basename(old_file)}: {e}")
            else:
                print(f"   • {os.path.basename(old_file)} - not found (OK)")
        
        # === VERIFY FILES ARE GONE ===
        print("\n🔍 Verifying cleanup...")
        if os.path.exists(extracted_text_path):
            print(f"   ⚠️  WARNING: {extracted_text_path} still exists!")
        else:
            print(f"   ✓ extracted_text.txt is gone")
            
        if os.path.exists(items_output_path):
            print(f"   ⚠️  WARNING: {items_output_path} still exists!")
        else:
            print(f"   ✓ items_offer1.json is gone")
        
        # === Step 1: Process with Document AI ===
        print("\n" + "=" * 60)
        print("🔍 STEP 1: Document AI Processing")
        print("=" * 60)
        
        test_process_path = os.path.join(BASE_DIR, 'test_process.py')
        result = subprocess.run(
            ['python', test_process_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=60
        )
        
        if result.returncode != 0:
            print(f"❌ Document AI Error")
            print(f"STDERR: {result.stderr}")
            return jsonify({
                'error': 'Document AI processing failed',
                'details': result.stderr
            }), 500
        
        print("✅ Document AI complete")
        print(result.stdout[-500:] if len(result.stdout) > 500 else result.stdout)
        
        # === Verify extracted text was created ===
        if not os.path.exists(extracted_text_path):
            print(f"❌ CRITICAL: Extracted text not created!")
            print(f"   Files in outputs: {os.listdir(OUTPUT_FOLDER)}")
            return jsonify({
                'error': 'Document AI did not create extracted text',
                'files_in_output': os.listdir(OUTPUT_FOLDER)
            }), 500
        
        text_size = os.path.getsize(extracted_text_path)
        print(f"✅ Extracted text created: {text_size} bytes")
        
        # === Step 2: Extract items with OpenAI ===
        print("\n" + "=" * 60)
        print("🤖 STEP 2: OpenAI Item Extraction")
        print("=" * 60)
        
        extract_items_path = os.path.join(BASE_DIR, 'extract_items.py')
        
        result = subprocess.run(
            ['python', extract_items_path, extracted_text_path, items_output_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=60
        )
        
        print(f"Return code: {result.returncode}")
        print(f"STDOUT:\n{result.stdout}")
        
        if result.stderr:
            print(f"STDERR:\n{result.stderr}")
        
        if result.returncode != 0:
            print(f"❌ OpenAI extraction failed")
            return jsonify({
                'error': 'Item extraction failed', 
                'details': result.stderr,
                'stdout': result.stdout
            }), 500
        
        print("✅ Extraction script completed")
        
        # === Verify items file was created ===
        if not os.path.exists(items_output_path):
            print(f"❌ CRITICAL: Items file not created!")
            print(f"   Files in outputs: {os.listdir(OUTPUT_FOLDER)}")
            return jsonify({
                'error': 'Items file not created',
                'files_in_output': os.listdir(OUTPUT_FOLDER)
            }), 500
        
        # === Step 3: Load and verify the NEW data ===
        print("\n" + "=" * 60)
        print("📊 STEP 3: Verifying Extracted Data")
        print("=" * 60)
        
        with open(items_output_path, 'r', encoding='utf-8') as f:
            file_content = f.read()
        
        print(f"File size: {len(file_content)} bytes")
        print(f"First 300 chars:\n{file_content[:300]}")
        
        full_data = json.loads(file_content)
        items = full_data.get('items', [])
        
        print(f"\n✅ EXTRACTION RESULTS:")
        print(f"   • Total items: {len(items)}")
        
        if len(items) > 0:
            print(f"   • First item name: {items[0].get('item_name', 'NO NAME')[:80]}")
            print(f"   • First item price: {items[0].get('unit_price', 'NO PRICE')}")
        
        # === CRITICAL CHECK: Verify this is NEW data ===
        if len(items) > 0:
            first_item_name = items[0].get('item_name', '').upper()
            if 'ISOBARIC' in first_item_name or 'MONOBLOCK' in first_item_name:
                print("\n" + "⚠️ " * 20)
                print("⚠️  WARNING: DETECTED OLD DATA!")
                print("⚠️  First item contains 'ISOBARIC MONOBLOCK'")
                print("⚠️  This means the cache was not cleared properly!")
                print("⚠️ " * 20)
        
        print("\n" + "=" * 60)
        print("✅ PROCESS OFFER 1 COMPLETE")
        print("=" * 60)
        
        return jsonify({
            'success': True,
            'items_count': len(items),
            'items': items,
            'full_data': full_data,
            'message': f'Successfully extracted {len(items)} items'
        })
        
    except subprocess.TimeoutExpired as e:
        print(f"❌ Timeout: {e}")
        return jsonify({'error': f'Processing timed out: {str(e)}'}), 500
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
        
        # Save as offer2_template.docx with absolute path
        filepath = os.path.join(BASE_DIR, 'offer2_template.docx')
        
        # Delete old template if exists
        if os.path.exists(filepath):
            os.remove(filepath)
            print(f"   ✓ Deleted old template")
        
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
    """
    Merge Offer 1 items + Offer 2 template
    Returns downloadable DOCX
    """
    try:
        print("=" * 60)
        print("🔄 GENERATE OFFER REQUEST")
        print("=" * 60)
        
        # Get optional parameters
        data = request.get_json() or {}
        markup = data.get('markup', 0)  # Percentage markup
        
        # Check if items exist with absolute path
        items_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        if not os.path.exists(items_path):
            return jsonify({'error': 'No items found. Please process Offer 1 first.'}), 400
        
        # === DEBUG: Show what's in the file ===
        print("\n📋 Checking items file...")
        with open(items_path, 'r', encoding='utf-8') as f:
            file_content = f.read()
        
        print(f"   File size: {len(file_content)} bytes")
        print(f"   First 500 chars:\n{file_content[:500]}")
        
        try:
            full_data = json.loads(file_content)
            items = full_data.get('items', [])
            print(f"   Items count: {len(items)}")
            if len(items) > 0:
                print(f"   First item: {json.dumps(items[0], ensure_ascii=False)[:200]}")
        except json.JSONDecodeError as e:
            print(f"   JSON ERROR: {e}")
            return jsonify({'error': f'Invalid JSON in items file: {str(e)}'}), 500
        
        # Apply markup if requested
        if markup > 0:
            print(f"\n💰 Applying {markup}% markup...")
            items = apply_markup_to_items(items, markup)
            full_data['items'] = items
            
            # Save updated items
            with open(items_path, 'w', encoding='utf-8') as f:
                json.dump(full_data, f, ensure_ascii=False, indent=2)
        
        # Run the generation script with absolute path
        print("\n📝 Running generation script...")
        generate_script_path = os.path.join(BASE_DIR, 'generate_offer_doc.py')
        
        result = subprocess.run(
            ['python', generate_script_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=30
        )
        
        print(f"\n🔍 Generation script results:")
        print(f"   Return code: {result.returncode}")
        print(f"   STDOUT length: {len(result.stdout)} chars")
        print(f"   STDERR length: {len(result.stderr)} chars")
        
        # Print full output for debugging
        if result.stdout:
            print(f"\n📤 STDOUT:\n{result.stdout}")
        
        if result.stderr:
            print(f"\n📤 STDERR:\n{result.stderr}")
        
        if result.returncode != 0:
            error_msg = result.stderr if result.stderr else "Unknown error"
            print(f"❌ Generation failed with code {result.returncode}")
            return jsonify({
                'error': 'Offer generation failed',
                'details': error_msg,
                'stdout': result.stdout,
                'return_code': result.returncode
            }), 500
        
        # Check if output file exists with absolute path
        output_path = os.path.join(OUTPUT_FOLDER, 'final_offer1.docx')
        if not os.path.exists(output_path):
            print("❌ Output file not created")
            return jsonify({
                'error': 'Output file not generated',
                'details': 'The generation script completed but no output file was created',
                'stdout': result.stdout,
                'stderr': result.stderr
            }), 500
        
        print(f"\n✅ Final offer generated successfully")
        print("=" * 60)
        
        return jsonify({
            'success': True,
            'message': 'Offer generated successfully',
            'download_url': '/api/download-offer',
            'items_count': len(items)
        })
        
    except subprocess.TimeoutExpired:
        print("❌ Generation timed out")
        return jsonify({'error': 'Generation timed out after 30 seconds'}), 500
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({
            'error': str(e),
            'traceback': traceback.format_exc()
        }), 500

# ============================================
# ENDPOINT 5: Download Generated Offer
# ============================================
@app.route('/api/download-offer', methods=['GET'])
def api_download_offer():
    """
    Download the generated final offer
    """
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
        # Try unit_price first
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
    print("🚀 Starting Requote AI Backend Server...")
    print(f"📡 Server will be available at: http://0.0.0.0:{port}")
    print("🌐 Ready to receive requests from Lovable frontend!\n")
    app.run(debug=True, host='0.0.0.0', port=port)