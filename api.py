from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import json
import subprocess
from werkzeug.utils import secure_filename
import threading
import time
import shutil

# Import the Python converter
from python_converter_final import convert_to_pdf_python

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

# Allowed file extensions - ALL formats now supported with Python converter!
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

def convert_to_docx_python(input_path, output_path, file_format):
    """Convert template formats to DOCX using Python"""
    try:
        print(f"Converting {file_format.upper()} template to DOCX...", flush=True)
        
        if file_format == 'docx':
            # Already DOCX, just copy
            shutil.copy(input_path, output_path)
            print("✓ DOCX template ready", flush=True)
            return True
        
        elif file_format == 'pdf':
            # PDF template - extract table structure and convert to DOCX
            import fitz  # PyMuPDF
            from docx import Document
            from docx.shared import Pt, RGBColor
            from docx.oxml.ns import qn
            from docx.oxml import OxmlElement
            
            print("Converting PDF template to DOCX with table extraction...", flush=True)
            
            # Read PDF
            pdf_doc = fitz.open(input_path)
            
            # Create new DOCX
            docx_doc = Document()
            
            # Process each page
            for page_num in range(len(pdf_doc)):
                page = pdf_doc[page_num]
                
                # Try to find tables in the PDF
                tables = page.find_tables()
                
                if tables:
                    print(f"  Found {len(tables)} table(s) on page {page_num + 1}", flush=True)
                    
                    for table_idx, table in enumerate(tables):
                        # Extract table data
                        table_data = table.extract()
                        
                        if not table_data or len(table_data) == 0:
                            continue
                        
                        # Create DOCX table
                        num_rows = len(table_data)
                        num_cols = max(len(row) for row in table_data) if table_data else 0
                        
                        if num_cols > 0:
                            docx_table = docx_doc.add_table(rows=num_rows, cols=num_cols)
                            docx_table.style = 'Light Grid Accent 1'
                            
                            # Fill table with data
                            for row_idx, row_data in enumerate(table_data):
                                for col_idx, cell_text in enumerate(row_data):
                                    if col_idx < num_cols:
                                        cell = docx_table.rows[row_idx].cells[col_idx]
                                        cell.text = str(cell_text) if cell_text else ""
                            
                            print(f"    ✓ Created table {table_idx + 1}: {num_rows} rows × {num_cols} cols", flush=True)
                else:
                    # No tables found, extract as text
                    text = page.get_text()
                    if text.strip():
                        # Try to detect if text looks like a table (has multiple columns)
                        lines = text.strip().split('\n')
                        
                        # Simple heuristic: if lines have consistent spacing, might be a table
                        if len(lines) > 2:
                            # Create a simple table from text
                            docx_table = docx_doc.add_table(rows=len(lines), cols=5)
                            docx_table.style = 'Light Grid Accent 1'
                            
                            for row_idx, line in enumerate(lines):
                                # Split line into cells (rough approximation)
                                parts = line.split()
                                for col_idx in range(min(5, len(parts))):
                                    cell = docx_table.rows[row_idx].cells[col_idx]
                                    cell.text = parts[col_idx] if col_idx < len(parts) else ""
                            
                            print(f"  Created text-based table on page {page_num + 1}", flush=True)
                        else:
                            para = docx_doc.add_paragraph(text)
                
                # Add page break except for last page
                if page_num < len(pdf_doc) - 1:
                    docx_doc.add_page_break()
            
            pdf_doc.close()
            
            # Save as DOCX
            docx_doc.save(output_path)
            print("✓ PDF converted to DOCX template with tables", flush=True)
            return True
        
        elif file_format == 'doc':
            # DOC files need LibreOffice, so we'll ask user to upload DOCX instead
            print("✗ DOC format requires LibreOffice. Please upload DOCX.", flush=True)
            return False
            
        elif file_format in ['xlsx', 'xls']:
            # Excel template - convert to DOCX preserving actual table structure
            import openpyxl
            from docx import Document
            from docx.shared import Pt, RGBColor
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            
            print(f"Converting {file_format.upper()} template to DOCX preserving table structure...", flush=True)
            
            # Read Excel
            workbook = openpyxl.load_workbook(input_path, data_only=True)
            sheet = workbook.active  # Use the first/active sheet
            
            # Create new DOCX
            docx_doc = Document()
            
            # Extract ALL rows from Excel
            all_rows = []
            for row in sheet.iter_rows(values_only=True):
                row_data = [str(cell) if cell is not None else '' for cell in row]
                all_rows.append(row_data)
            
            print(f"  Found {len(all_rows)} rows in Excel", flush=True)
            
            # Find the pricing table header (the row with keywords)
            table_start_row = None
            for idx, row in enumerate(all_rows):
                row_text = ' '.join(row).upper()
                if any(keyword in row_text for keyword in ['POSITION', 'DESCRIPTION', 'PRICE', 'QUANTITY', 'TOTAL']):
                    table_start_row = idx
                    print(f"  Found pricing table header at row {idx + 1}", flush=True)
                    break
            
            if table_start_row is None:
                # No clear table found, use generic approach
                print(f"  No pricing table header found, using generic conversion", flush=True)
                table_start_row = 0
            
            # Add header rows (before the table) as paragraphs
            for idx in range(table_start_row):
                row_text = ' '.join(all_rows[idx]).strip()
                if row_text:
                    para = docx_doc.add_paragraph(row_text)
                    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            if table_start_row > 0:
                docx_doc.add_paragraph()  # Spacing
            
            # Create DOCX table from Excel data (from table header onwards)
            table_rows = all_rows[table_start_row:]
            
            if table_rows:
                # Determine number of columns (max columns in any row)
                num_cols = max(len(row) for row in table_rows)
                num_rows = len(table_rows)
                
                print(f"  Creating table: {num_rows} rows × {num_cols} columns", flush=True)
                
                # Create table
                docx_table = docx_doc.add_table(rows=num_rows, cols=num_cols)
                docx_table.style = 'Light Grid Accent 1'
                
                # Fill table with Excel data
                for row_idx, row_data in enumerate(table_rows):
                    for col_idx in range(num_cols):
                        cell_text = row_data[col_idx] if col_idx < len(row_data) else ''
                        docx_table.rows[row_idx].cells[col_idx].text = cell_text
                        
                        # Make first row bold (header)
                        if row_idx == 0:
                            for paragraph in docx_table.rows[row_idx].cells[col_idx].paragraphs:
                                for run in paragraph.runs:
                                    run.font.bold = True
                
                print(f"  ✓ Table created with all Excel data", flush=True)
            
            # Save as DOCX
            docx_doc.save(output_path)
            print(f"✓ {file_format.upper()} converted to DOCX template", flush=True)
            return True
            
        else:
            print(f"✗ Unsupported template format: {file_format}", flush=True)
            return False
            
    except Exception as e:
        print(f"✗ Template conversion error: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
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
        'version': 'SV10-PythonConverter',
        'status': 'healthy',
        'conversion_method': 'Pure Python (no LibreOffice required)',
        'supported_formats': {
            'offer1': list(ALLOWED_OFFER1_EXTENSIONS),
            'offer2': list(ALLOWED_OFFER2_EXTENSIONS)
        }
    })

def process_file_background(filepath, file_extension):
    """Background processing for any supported file format"""
    global processing_status
    
    try:
        with _status_lock:
            processing_status['status'] = 'processing'
            processing_status['message'] = f'Processing {file_extension.upper()} file...'
            processing_status['file_format'] = file_extension
            processing_status['started_at'] = time.time()
            processing_status['updated_at'] = time.time()
        
        print("=== BACKGROUND PROCESSING STARTED ===", flush=True)
        print(f"Format: {file_extension}", flush=True)
        print(f"File: {filepath}", flush=True)
        
        # Determine processing path based on format
        pdf_path = os.path.join(UPLOAD_FOLDER, 'offer1.pdf')
        
        if file_extension == 'pdf':
            # Already PDF, just copy
            if filepath != pdf_path:
                shutil.copy(filepath, pdf_path)
            print("✓ PDF ready for processing", flush=True)
            
        elif file_extension in ['docx', 'doc', 'xlsx', 'xls', 'png', 'jpg', 'jpeg']:
            # Convert to PDF using Python (NO LibreOffice needed!)
            with _status_lock:
                processing_status['message'] = f'Converting {file_extension.upper()} to PDF using Python...'
                processing_status['updated_at'] = time.time()
            
            conversion_success = convert_to_pdf_python(filepath, pdf_path, file_extension)
            
            if not conversion_success:
                with _status_lock:
                    processing_status['status'] = 'error'
                    processing_status['message'] = f'Failed to convert {file_extension.upper()} to PDF. Please try a different file or contact support.'
                    processing_status['updated_at'] = time.time()
                return
        else:
            with _status_lock:
                processing_status['status'] = 'error'
                processing_status['message'] = f'Unsupported file format: {file_extension}'
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
            processing_status['message'] = f'Successfully extracted {len(items)} items from {file_extension.upper()}'
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
    """Process Offer 1 - supports multiple formats with Python converter"""
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
    """Upload Offer 2 template - DOCX only (other formats not suitable for templates)"""
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
        
        # Save template
        template_path = os.path.join(BASE_DIR, 'offer2_template.docx')
        
        if file_extension == 'docx':
            # DOCX is perfect - just save it
            file.save(template_path)
            print(f"✓ DOCX template saved directly", flush=True)
        else:
            # Other formats - save original and try to convert
            original_path = os.path.join(BASE_DIR, f'offer2_template_original.{file_extension}')
            file.save(original_path)
            
            print(f"Converting {file_extension.upper()} template to DOCX...", flush=True)
            print(f"  Original file: {original_path} ({os.path.getsize(original_path)} bytes)", flush=True)
            
            conversion_success = convert_to_docx_python(original_path, template_path, file_extension)
            
            if not conversion_success:
                print(f"✗ Conversion returned False", flush=True)
                return jsonify({
                    'error': f'Cannot convert {file_extension.upper()} to template format.',
                    'suggestion': 'Please try uploading a DOCX template instead.'
                }), 500
            
            # Verify the file was actually created
            if not os.path.exists(template_path):
                print(f"✗ Template file not created at {template_path}", flush=True)
                return jsonify({
                    'error': f'Conversion failed - output file not created.',
                    'suggestion': 'Please try uploading a DOCX template instead.'
                }), 500
            
            # Verify it's a valid DOCX
            try:
                from docx import Document
                test_doc = Document(template_path)
                print(f"✓ Template validation: {len(test_doc.tables)} tables, {len(test_doc.paragraphs)} paragraphs", flush=True)
            except Exception as ve:
                print(f"✗ Template validation failed: {str(ve)}", flush=True)
                return jsonify({
                    'error': f'Converted template is invalid.',
                    'suggestion': 'Please try uploading a DOCX template instead.'
                }), 500
        
        if os.path.exists(template_path):
            print(f"✓ Template ready: {template_path}", flush=True)
            print(f"✓ Size: {os.path.getsize(template_path)} bytes", flush=True)
        print("=" * 60, flush=True)
        
        return jsonify({
            'success': True,
            'message': f'Template uploaded successfully ({file_extension.upper()})',
            'file_format': file_extension
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
    print("Starting Requote AI Backend - SV10 Python Converter")
    print(f"Server at: http://0.0.0.0:{port}")
    print(f"Conversion method: Pure Python (no LibreOffice)")
    print(f"Supported Offer 1 formats: {', '.join(ALLOWED_OFFER1_EXTENSIONS)}")
    print(f"Supported Offer 2 formats: {', '.join(ALLOWED_OFFER2_EXTENSIONS)}")
    print("=" * 60)
    app.run(debug=True, host='0.0.0.0', port=port)