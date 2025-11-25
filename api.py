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
    'file_format': None,
    'system': 'sv12'  # 'sv12' or 'flexible'
}

# Template cache for flexible system
template_structure_cache = {}

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
            shutil.copy(input_path, output_path)
            print("✓ DOCX template ready", flush=True)
            return True
        
        elif file_format == 'pdf':
            import fitz
            from docx import Document
            from docx.shared import Pt, RGBColor
            
            print("Converting PDF template to DOCX with table extraction...", flush=True)
            pdf_doc = fitz.open(input_path)
            docx_doc = Document()
            
            for page_num in range(len(pdf_doc)):
                page = pdf_doc[page_num]
                tables = page.find_tables()
                
                if tables:
                    print(f"  Found {len(tables)} table(s) on page {page_num + 1}", flush=True)
                    for table in tables:
                        table_data = table.extract()
                        if not table_data or len(table_data) == 0:
                            continue
                        
                        num_rows = len(table_data)
                        num_cols = max(len(row) for row in table_data) if table_data else 0
                        
                        if num_cols > 0:
                            docx_table = docx_doc.add_table(rows=num_rows, cols=num_cols)
                            docx_table.style = 'Light Grid Accent 1'
                            
                            for row_idx, row_data in enumerate(table_data):
                                for col_idx, cell_text in enumerate(row_data):
                                    if col_idx < num_cols:
                                        cell = docx_table.rows[row_idx].cells[col_idx]
                                        cell.text = str(cell_text) if cell_text else ""
                else:
                    text = page.get_text()
                    if text.strip():
                        para = docx_doc.add_paragraph(text)
                
                if page_num < len(pdf_doc) - 1:
                    docx_doc.add_page_break()
            
            pdf_doc.close()
            docx_doc.save(output_path)
            print("✓ PDF converted to DOCX template with tables", flush=True)
            return True
        
        elif file_format == 'doc':
            print("✗ DOC format requires LibreOffice. Please upload DOCX.", flush=True)
            return False
            
        elif file_format in ['xlsx', 'xls']:
            import openpyxl
            from docx import Document
            from docx.shared import Pt
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            
            print(f"Converting {file_format.upper()} template to DOCX...", flush=True)
            workbook = openpyxl.load_workbook(input_path, data_only=True)
            sheet = workbook.active
            docx_doc = Document()
            
            all_rows = []
            for row in sheet.iter_rows(values_only=True):
                row_data = [str(cell) if cell is not None else '' for cell in row]
                all_rows.append(row_data)
            
            print(f"  Found {len(all_rows)} rows in Excel", flush=True)
            
            table_start_row = None
            for idx, row in enumerate(all_rows):
                row_text = ' '.join(row).upper()
                if any(keyword in row_text for keyword in ['POSITION', 'DESCRIPTION', 'PRICE', 'QUANTITY', 'TOTAL']):
                    table_start_row = idx
                    break
            
            if table_start_row is None:
                table_start_row = 0
            
            for idx in range(table_start_row):
                row_text = ' '.join(all_rows[idx]).strip()
                if row_text:
                    para = docx_doc.add_paragraph(row_text)
            
            if table_start_row > 0:
                docx_doc.add_paragraph()
            
            table_rows = all_rows[table_start_row:]
            
            if table_rows:
                num_cols = max(len(row) for row in table_rows)
                num_rows = len(table_rows)
                
                docx_table = docx_doc.add_table(rows=num_rows, cols=num_cols)
                docx_table.style = 'Light Grid Accent 1'
                
                for row_idx, row_data in enumerate(table_rows):
                    for col_idx in range(num_cols):
                        cell_text = row_data[col_idx] if col_idx < len(row_data) else ''
                        docx_table.rows[row_idx].cells[col_idx].text = cell_text
                        
                        if row_idx == 0:
                            for paragraph in docx_table.rows[row_idx].cells[col_idx].paragraphs:
                                for run in paragraph.runs:
                                    run.font.bold = True
            
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
        'version': 'SV12-Flexible-Hybrid',
        'status': 'healthy',
        'systems': {
            'sv12': 'Stable hardcoded system (backup)',
            'flexible': 'New GPT-driven 3-prompt system'
        },
        'supported_formats': {
            'offer1': list(ALLOWED_OFFER1_EXTENSIONS),
            'offer2': list(ALLOWED_OFFER2_EXTENSIONS)
        }
    })

def process_file_background_sv12(filepath, file_extension):
    """Background processing using SV12 (original hardcoded system)"""
    global processing_status
    
    try:
        with _status_lock:
            processing_status['status'] = 'processing'
            processing_status['message'] = f'Processing {file_extension.upper()} file with SV12...'
            processing_status['file_format'] = file_extension
            processing_status['system'] = 'sv12'
            processing_status['started_at'] = time.time()
            processing_status['updated_at'] = time.time()
        
        print("=== SV12 BACKGROUND PROCESSING STARTED ===", flush=True)
        
        pdf_path = os.path.join(UPLOAD_FOLDER, 'offer1.pdf')
        
        if file_extension == 'pdf':
            if filepath != pdf_path:
                shutil.copy(filepath, pdf_path)
        else:
            with _status_lock:
                processing_status['message'] = f'Converting {file_extension.upper()} to PDF...'
            
            conversion_success = convert_to_pdf_python(filepath, pdf_path, file_extension)
            if not conversion_success:
                with _status_lock:
                    processing_status['status'] = 'error'
                    processing_status['message'] = f'Failed to convert {file_extension.upper()}'
                return
        
        with _status_lock:
            processing_status['message'] = 'Extracting items (SV12)...'
        
        items_output_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        extract_script_path = os.path.join(BASE_DIR, 'extract_pdf_direct_enhanced.py')
        
        result = subprocess.run(
            ['python', extract_script_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=300
        )
        
        if result.returncode != 0 or not os.path.exists(items_output_path):
            with _status_lock:
                processing_status['status'] = 'error'
                processing_status['message'] = 'Extraction failed (SV12)'
            return
        
        with open(items_output_path, 'r', encoding='utf-8') as f:
            full_data = json.load(f)
        
        items = full_data.get('items', [])
        
        with _status_lock:
            processing_status['status'] = 'completed'
            processing_status['message'] = f'Successfully extracted {len(items)} items (SV12)'
            processing_status['items_count'] = len(items)
            processing_status['items'] = items
            processing_status['updated_at'] = time.time()
        
    except Exception as e:
        print(f"=== SV12 ERROR: {str(e)} ===", flush=True)
        with _status_lock:
            processing_status['status'] = 'error'
            processing_status['message'] = f'SV12 error: {str(e)}'

def process_file_background_flexible(filepath, file_extension):
    """Background processing using FLEXIBLE 3-PROMPT SYSTEM"""
    global processing_status
    
    try:
        with _status_lock:
            processing_status['status'] = 'processing'
            processing_status['message'] = f'Processing {file_extension.upper()} file with Flexible System...'
            processing_status['file_format'] = file_extension
            processing_status['system'] = 'flexible'
            processing_status['started_at'] = time.time()
            processing_status['updated_at'] = time.time()
        
        print("=== FLEXIBLE SYSTEM BACKGROUND PROCESSING STARTED ===", flush=True)
        
        # STEP 1: Convert to PDF if needed
        pdf_path = os.path.join(UPLOAD_FOLDER, 'offer1.pdf')
        
        if file_extension == 'pdf':
            if filepath != pdf_path:
                shutil.copy(filepath, pdf_path)
        else:
            with _status_lock:
                processing_status['message'] = f'Converting {file_extension.upper()} to PDF...'
            
            conversion_success = convert_to_pdf_python(filepath, pdf_path, file_extension)
            if not conversion_success:
                with _status_lock:
                    processing_status['status'] = 'error'
                    processing_status['message'] = f'Failed to convert {file_extension.upper()}'
                return
        
        # STEP 2: Extract using PROMPT 1 (extract_pdf_direct_enhanced_v2.py)
        with _status_lock:
            processing_status['message'] = 'Extracting with PROMPT 1 (Flexible)...'
        
        extract_script_path = os.path.join(BASE_DIR, 'extract_pdf_direct_enhanced_v2.py')
        
        print("Running PROMPT 1 extraction...", flush=True)
        result = subprocess.run(
            ['python', extract_script_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=300
        )
        
        if result.stdout:
            print(result.stdout, flush=True)
        if result.stderr:
            print(result.stderr, flush=True)
        
        items_output_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        
        if result.returncode != 0 or not os.path.exists(items_output_path):
            with _status_lock:
                processing_status['status'] = 'error'
                processing_status['message'] = 'PROMPT 1 extraction failed'
            return
        
        with open(items_output_path, 'r', encoding='utf-8') as f:
            full_data = json.load(f)
        
        items = full_data.get('items', [])
        
        print(f"✓ Extracted {len(items)} items with Flexible System", flush=True)
        
        with _status_lock:
            processing_status['status'] = 'completed'
            processing_status['message'] = f'Successfully extracted {len(items)} items (Flexible)'
            processing_status['items_count'] = len(items)
            processing_status['items'] = items
            processing_status['updated_at'] = time.time()
        
    except Exception as e:
        print(f"=== FLEXIBLE SYSTEM ERROR: {str(e)} ===", flush=True)
        import traceback
        traceback.print_exc()
        with _status_lock:
            processing_status['status'] = 'error'
            processing_status['message'] = f'Flexible system error: {str(e)}'

@app.route('/api/process-offer1', methods=['POST', 'OPTIONS'])
def api_process_offer1():
    """Process Offer 1 - supports both SV12 and Flexible systems"""
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
        
        if not allowed_file(file.filename, ALLOWED_OFFER1_EXTENSIONS):
            return jsonify({
                'error': f'Unsupported file format. Allowed: {", ".join(ALLOWED_OFFER1_EXTENSIONS)}'
            }), 400
        
        # Get system preference (default: sv12 for stability)
        use_flexible = request.form.get('use_flexible', 'false').lower() == 'true'
        
        file_extension = get_file_extension(file.filename)
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, f'offer1_original.{file_extension}')
        file.save(filepath)
        
        print(f"✓ File saved: {filepath}", flush=True)
        print(f"✓ Format: {file_extension.upper()}", flush=True)
        print(f"✓ System: {'FLEXIBLE' if use_flexible else 'SV12'}", flush=True)
        
        # Reset status
        with _status_lock:
            processing_status = {
                'status': 'processing',
                'message': f'File uploaded, starting processing...',
                'items_count': 0,
                'items': [],
                'file_format': file_extension,
                'system': 'flexible' if use_flexible else 'sv12',
                'started_at': time.time(),
                'updated_at': time.time()
            }
        
        # Start background thread with chosen system
        if use_flexible:
            thread = threading.Thread(
                target=process_file_background_flexible,
                args=(filepath, file_extension),
                daemon=True
            )
        else:
            thread = threading.Thread(
                target=process_file_background_sv12,
                args=(filepath, file_extension),
                daemon=True
            )
        
        thread.start()
        
        print(f"✓ Background thread started ({'FLEXIBLE' if use_flexible else 'SV12'})", flush=True)
        print("=" * 60, flush=True)
        
        return jsonify({
            'success': True,
            'message': f'Processing with {"Flexible" if use_flexible else "SV12"} system. Poll /api/status for updates.',
            'status': 'processing',
            'file_format': file_extension,
            'system': 'flexible' if use_flexible else 'sv12'
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
    
    if status_copy.get('started_at') and status_copy['status'] == 'processing':
        elapsed = time.time() - status_copy['started_at']
        status_copy['elapsed_seconds'] = round(elapsed, 1)
    
    return jsonify(status_copy)

@app.route('/api/analyze-template', methods=['POST', 'OPTIONS'])
def api_analyze_template():
    """NEW: Analyze Offer 2 template using PROMPT 2 (Flexible system only)"""
    if request.method == 'OPTIONS':
        return '', 204
    
    try:
        print("=" * 60, flush=True)
        print("Analyzing template with PROMPT 2", flush=True)
        
        # Run analyze_offer2_template.py
        analyze_script_path = os.path.join(BASE_DIR, 'analyze_offer2_template.py')
        
        result = subprocess.run(
            ['python', analyze_script_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=120
        )
        
        if result.stdout:
            print(result.stdout, flush=True)
        if result.stderr:
            print(result.stderr, flush=True)
        
        template_structure_path = os.path.join(OUTPUT_FOLDER, 'template_structure.json')
        
        if result.returncode != 0 or not os.path.exists(template_structure_path):
            return jsonify({
                'error': 'Template analysis failed',
                'details': result.stderr
            }), 500
        
        with open(template_structure_path, 'r', encoding='utf-8') as f:
            structure = json.load(f)
        
        print("✓ Template analyzed successfully", flush=True)
        
        return jsonify({
            'success': True,
            'message': 'Template analyzed with PROMPT 2',
            'structure': structure
        })
        
    except Exception as e:
        print(f"ERROR in analyze-template: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/upload-offer2', methods=['POST', 'OPTIONS'])
def api_upload_offer2():
    """Upload Offer 2 template - DOCX or XLSX"""
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
        
        if not allowed_file(file.filename, ALLOWED_OFFER2_EXTENSIONS):
            return jsonify({
                'error': f'Unsupported template format. Allowed: {", ".join(ALLOWED_OFFER2_EXTENSIONS)}'
            }), 400
        
        file_extension = get_file_extension(file.filename)
        
        # Get system preference
        use_flexible = request.form.get('use_flexible', 'false').lower() == 'true'
        
        print(f"✓ Template format: {file_extension.upper()}", flush=True)
        print(f"✓ System: {'FLEXIBLE' if use_flexible else 'SV12'}", flush=True)
        
        # Clean up old template files
        old_docx = os.path.join(BASE_DIR, 'offer2_template.docx')
        old_xlsx = os.path.join(BASE_DIR, 'offer2_template.xlsx')
        old_xls = os.path.join(BASE_DIR, 'offer2_template.xls')
        old_format = os.path.join(BASE_DIR, 'template_format.txt')
        
        for old_file in [old_docx, old_xlsx, old_xls, old_format]:
            if os.path.exists(old_file):
                os.remove(old_file)
        
        # Save template
        if file_extension in ['xlsx', 'xls']:
            template_path = os.path.join(BASE_DIR, f'offer2_template.{file_extension}')
            file.save(template_path)
            with open(os.path.join(BASE_DIR, 'template_format.txt'), 'w') as f:
                f.write('xlsx')
            
        elif file_extension == 'docx':
            template_path = os.path.join(BASE_DIR, 'offer2_template.docx')
            file.save(template_path)
            with open(os.path.join(BASE_DIR, 'template_format.txt'), 'w') as f:
                f.write('docx')
        else:
            # Other formats - convert
            template_path = os.path.join(BASE_DIR, 'offer2_template.docx')
            original_path = os.path.join(BASE_DIR, f'offer2_template_original.{file_extension}')
            file.save(original_path)
            
            conversion_success = convert_to_docx_python(original_path, template_path, file_extension)
            
            if not conversion_success:
                return jsonify({
                    'error': f'Cannot convert {file_extension.upper()} to template format.',
                    'suggestion': 'Please try uploading a DOCX template instead.'
                }), 500
        
        # If using flexible system, analyze template immediately
        if use_flexible:
            print("Analyzing template with PROMPT 2...", flush=True)
            analyze_script_path = os.path.join(BASE_DIR, 'analyze_offer2_template.py')
            
            result = subprocess.run(
                ['python', analyze_script_path],
                capture_output=True,
                text=True,
                cwd=BASE_DIR,
                timeout=120
            )
            
            if result.returncode == 0:
                print("✓ Template analyzed", flush=True)
            else:
                print("⚠ Template analysis failed, will retry during generation", flush=True)
        
        print(f"✓ Template ready: {template_path}", flush=True)
        print("=" * 60, flush=True)
        
        return jsonify({
            'success': True,
            'message': f'Template uploaded successfully ({file_extension.upper()})',
            'file_format': file_extension,
            'analyzed': use_flexible
        })
        
    except Exception as e:
        print(f"Error in upload-offer2: {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/generate-offer', methods=['POST', 'OPTIONS'])
def api_generate_offer():
    """Generate offer - supports both SV12 and Flexible systems"""
    if request.method == 'OPTIONS':
        return '', 204
    
    try:
        print("Starting offer generation...", flush=True)
        
        data = request.get_json() or {}
        markup = data.get('markup', 0)
        use_flexible = data.get('use_flexible', False)
        
        print(f"System: {'FLEXIBLE' if use_flexible else 'SV12'}", flush=True)
        
        items_path = os.path.join(OUTPUT_FOLDER, 'items_offer1.json')
        if not os.path.exists(items_path):
            return jsonify({'error': 'No items found. Please process Offer 1 first.'}), 400
        
        # Apply markup if needed
        if markup > 0:
            print(f"Applying {markup}% markup...", flush=True)
            with open(items_path, 'r', encoding='utf-8') as f:
                full_data = json.load(f)
            
            items = full_data.get('items', [])
            items = apply_markup_to_items(items, markup)
            full_data['items'] = items
            
            with open(items_path, 'w', encoding='utf-8') as f:
                json.dump(full_data, f, ensure_ascii=False, indent=2)
        
        # Determine template format
        template_docx = os.path.join(BASE_DIR, 'offer2_template.docx')
        template_xlsx = os.path.join(BASE_DIR, 'offer2_template.xlsx')
        
        if os.path.exists(template_docx):
            template_format = 'docx'
        elif os.path.exists(template_xlsx):
            template_format = 'xlsx'
        else:
            return jsonify({'error': 'No template found. Please upload Offer 2 template first.'}), 400
        
        # Clean up old output files
        old_docx = os.path.join(OUTPUT_FOLDER, 'final_offer1.docx')
        old_xlsx = os.path.join(OUTPUT_FOLDER, 'final_offer1.xlsx')
        if os.path.exists(old_docx):
            os.remove(old_docx)
        if os.path.exists(old_xlsx):
            os.remove(old_xlsx)
        
        # Generate using chosen system
        if use_flexible:
            print("Generating with FLEXIBLE system (PROMPT 3)...", flush=True)
            generate_script_path = os.path.join(BASE_DIR, 'generate_offer_flexible.py')
        else:
            print("Generating with SV12 system...", flush=True)
            if template_format == 'xlsx':
                generate_script_path = os.path.join(BASE_DIR, 'generate_offer_xlsx.py')
            else:
                generate_script_path = os.path.join(BASE_DIR, 'generate_offer_doc.py')
        
        result = subprocess.run(
            ['python', generate_script_path],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=120
        )
        
        if result.stdout:
            print(result.stdout, flush=True)
        if result.stderr:
            print(result.stderr, flush=True)
        
        if result.returncode != 0:
            return jsonify({
                'error': 'Offer generation failed',
                'details': result.stderr
            }), 500
        
        output_filename = f'final_offer1.{template_format}'
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        
        if not os.path.exists(output_path):
            return jsonify({'error': 'Output file not generated'}), 500
        
        with open(items_path, 'r', encoding='utf-8') as f:
            full_data = json.load(f)
            items = full_data.get('items', [])
        
        print(f"✓ Offer generated successfully with {'FLEXIBLE' if use_flexible else 'SV12'}", flush=True)
        
        return jsonify({
            'success': True,
            'message': 'Offer generated successfully',
            'download_url': '/api/download-offer',
            'items_count': len(items),
            'output_format': template_format,
            'system': 'flexible' if use_flexible else 'sv12'
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
        output_xlsx = os.path.join(OUTPUT_FOLDER, 'final_offer1.xlsx')
        output_docx = os.path.join(OUTPUT_FOLDER, 'final_offer1.docx')
        
        if os.path.exists(output_xlsx):
            output_path = output_xlsx
            mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            download_name = 'requoted_offer.xlsx'
        elif os.path.exists(output_docx):
            output_path = output_docx
            mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            download_name = 'requoted_offer.docx'
        else:
            return jsonify({'error': 'No offer generated yet'}), 404
        
        from flask import Response
        
        with open(output_path, 'rb') as f:
            file_data = f.read()
        
        response = Response(
            file_data,
            mimetype=mimetype,
            headers={
                'Content-Disposition': f'attachment; filename="{download_name}"',
                'Content-Type': mimetype,
                'Cache-Control': 'no-cache',
                'Pragma': 'no-cache',
                'Expires': '0'
            }
        )
        
        return response
        
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
    print("Starting Requote AI Backend - SV12 + Flexible Hybrid")
    print(f"Server at: http://0.0.0.0:{port}")
    print(f"Systems available:")
    print(f"  - SV12: Stable hardcoded system (default)")
    print(f"  - Flexible: New GPT-driven 3-prompt system")
    print(f"Supported formats: {', '.join(ALLOWED_OFFER1_EXTENSIONS)}")
    print("=" * 60)
    app.run(debug=True, host='0.0.0.0', port=port)