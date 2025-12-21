"""
Standard Offer 3 Template Structure
Defines the professional quotation format that will be generated
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import base64
import io
from PIL import Image as PILImage

class Offer3Template:
    """
    Standard professional quotation template
    Creates consistent, high-quality output regardless of input formats
    """
    
    def __init__(self):
        self.doc = Document()
        
        # Standard styling
        self.font_name = "Calibri"
        self.font_size_body = 11
        self.font_size_header = 14
        self.font_size_small = 9
        
        # Standard colors
        self.color_header_bg = "4472C4"  # Professional blue
        self.color_text_dark = "000000"
        self.color_text_light = "666666"
        
    def add_header_section(self, company_data):
        """
        Add company header with logo and contact information
        Replicates the exact header from Offer 2
        """
        
        print("\n" + "=" * 60, flush=True)
        print("ADDING HEADER SECTION", flush=True)
        print("=" * 60, flush=True)
        
        # Debug: Show what we received
        print(f"Company name: {company_data.get('company_name', 'MISSING')}", flush=True)
        print(f"Address: {company_data.get('address', 'MISSING')[:50] if company_data.get('address') else 'MISSING'}...", flush=True)
        print(f"Logo present: {bool(company_data.get('logo'))}", flush=True)
        
        # Add logo if available
        if company_data.get('logo') and company_data['logo'].get('data'):
            try:
                logo_format = company_data['logo']['format']
                logo_base64 = company_data['logo']['data']
                
                # Decode base64 to bytes
                logo_bytes = base64.b64decode(logo_base64)
                
                # Save to temporary file
                temp_logo_path = '/tmp/temp_logo.' + logo_format
                with open(temp_logo_path, 'wb') as f:
                    f.write(logo_bytes)
                
                # Add logo to document
                logo_para = self.doc.add_paragraph()
                logo_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = logo_para.add_run()
                
                # Add logo with original sizing (max 2.5 inches wide)
                run.add_picture(temp_logo_path, width=Inches(2.5))
                
                print(f"✓ Logo added to header ({len(logo_bytes)} bytes, {logo_format})", flush=True)
                
                # Clean up temp file
                try:
                    os.remove(temp_logo_path)
                except:
                    pass
                
            except Exception as e:
                print(f"⚠ Could not add logo: {str(e)}", flush=True)
        else:
            print("⚠ No logo data available", flush=True)
        
        # Company information section
        # Company name - ALWAYS show, use placeholder if missing
        company_name = company_data.get('company_name', '[COMPANY NAME]')
        if company_name and company_name.strip():
            p = self.doc.add_paragraph()
            run = p.add_run(company_name)
            run.font.size = Pt(self.font_size_header)
            run.font.bold = True
            run.font.name = self.font_name
            print(f"✓ Added company name: {company_name}", flush=True)
        else:
            print("✗ Company name missing - using placeholder", flush=True)
            p = self.doc.add_paragraph()
            run = p.add_run('[COMPANY NAME - NOT EXTRACTED]')
            run.font.size = Pt(self.font_size_header)
            run.font.bold = True
            run.font.name = self.font_name
        
        # Address
        address = company_data.get('address', '')
        if address and address.strip():
            p = self.doc.add_paragraph()
            run = p.add_run(address)
            run.font.size = Pt(self.font_size_body)
            run.font.name = self.font_name
            print(f"✓ Added address", flush=True)
        else:
            print("⚠ No address data", flush=True)
        
        # Contact details
        contact_parts = []
        if company_data.get('phone'):
            contact_parts.append(f"Phone: {company_data['phone']}")
        if company_data.get('email'):
            contact_parts.append(f"Email: {company_data['email']}")
        if company_data.get('website'):
            contact_parts.append(f"Website: {company_data['website']}")
        
        if contact_parts:
            p = self.doc.add_paragraph()
            run = p.add_run(" | ".join(contact_parts))
            run.font.size = Pt(self.font_size_body)
            run.font.name = self.font_name
            print(f"✓ Added contact details: {len(contact_parts)} fields", flush=True)
        else:
            print("⚠ No contact details", flush=True)
        
        # Spacing after header
        self.doc.add_paragraph()
        print("=" * 60, flush=True)
    
    def add_document_info_table(self, quote_number, date, valid_until, customer_name):
        """
        Add document metadata table:
        - Quotation No
        - Date
        - Valid Until
        - Customer
        """
        
        info_table = self.doc.add_table(rows=4, cols=2)
        info_table.style = 'Light Grid Accent 1'
        
        # Set column widths
        for row in info_table.rows:
            row.cells[0].width = Inches(2.0)
            row.cells[1].width = Inches(4.0)
        
        # Row 1: Quotation Number
        self._set_cell_text(info_table.rows[0].cells[0], "Quotation No:", bold=True)
        self._set_cell_text(info_table.rows[0].cells[1], quote_number)
        
        # Row 2: Date
        self._set_cell_text(info_table.rows[1].cells[0], "Date:", bold=True)
        self._set_cell_text(info_table.rows[1].cells[1], date)
        
        # Row 3: Valid Until
        self._set_cell_text(info_table.rows[2].cells[0], "Valid Until:", bold=True)
        self._set_cell_text(info_table.rows[2].cells[1], valid_until)
        
        # Row 4: Customer
        self._set_cell_text(info_table.rows[3].cells[0], "Customer:", bold=True)
        self._set_cell_text(info_table.rows[3].cells[1], customer_name)
        
        # Spacing after table
        self.doc.add_paragraph()
    
    def add_pricing_table(self, items, currency="€"):
        """
        Add pricing table with all items grouped by category
        
        Columns: No. | Description | Quantity | Unit Price | Total
        """
        
        # Create table with header row
        pricing_table = self.doc.add_table(rows=1, cols=5)
        pricing_table.style = 'Light Grid Accent 1'
        
        # Set column widths
        pricing_table.columns[0].width = Inches(0.5)   # No.
        pricing_table.columns[1].width = Inches(3.5)   # Description
        pricing_table.columns[2].width = Inches(0.8)   # Quantity
        pricing_table.columns[3].width = Inches(1.0)   # Unit Price
        pricing_table.columns[4].width = Inches(1.0)   # Total
        
        # Header row
        header_cells = pricing_table.rows[0].cells
        self._set_cell_text(header_cells[0], "No.", bold=True, bg_color=self.color_header_bg, text_color="FFFFFF")
        self._set_cell_text(header_cells[1], "Description", bold=True, bg_color=self.color_header_bg, text_color="FFFFFF")
        self._set_cell_text(header_cells[2], "Quantity", bold=True, bg_color=self.color_header_bg, text_color="FFFFFF")
        self._set_cell_text(header_cells[3], "Unit Price", bold=True, bg_color=self.color_header_bg, text_color="FFFFFF")
        self._set_cell_text(header_cells[4], "Total", bold=True, bg_color=self.color_header_bg, text_color="FFFFFF")
        
        # Group items by category
        from collections import OrderedDict
        categorized = OrderedDict()
        for item in items:
            cat = item.get('category', 'Items')
            if cat not in categorized:
                categorized[cat] = []
            categorized[cat].append(item)
        
        item_counter = 1
        
        # Add items by category
        for category, cat_items in categorized.items():
            # Category header row
            cat_row = pricing_table.add_row().cells
            cat_row[0].merge(cat_row[4])
            self._set_cell_text(cat_row[0], category, bold=True, bg_color="E7E6E6")
            
            # Items in category
            for item in cat_items:
                row = pricing_table.add_row().cells
                
                # No.
                self._set_cell_text(row[0], str(item_counter), align="center")
                item_counter += 1
                
                # Description (item_name only, technical details go to separate section)
                description = item.get('item_name', '')
                self._set_cell_text(row[1], description)
                
                # Quantity
                qty = item.get('quantity', '1')
                self._set_cell_text(row[2], str(qty), align="center")
                
                # Unit Price
                unit_price = item.get('unit_price', '')
                self._set_cell_text(row[3], str(unit_price), align="right")
                
                # Total
                total = item.get('total_price', item.get('unit_price', ''))
                self._set_cell_text(row[4], str(total), align="right")
        
        # Add totals row if needed (placeholder for now)
        totals_row = pricing_table.add_row().cells
        totals_row[0].merge(totals_row[3])
        self._set_cell_text(totals_row[0], "TOTAL:", bold=True, align="right")
        self._set_cell_text(totals_row[4], "", bold=True, align="right")
        
        # Spacing after table
        self.doc.add_paragraph()
    
    def add_technical_descriptions(self, items):
        """
        Add detailed technical descriptions for each item
        Only for items that have description/specifications
        """
        
        # Section heading
        heading = self.doc.add_paragraph()
        run = heading.add_run("Technical Specifications")
        run.font.size = Pt(self.font_size_header)
        run.font.bold = True
        run.font.name = self.font_name
        heading.style = 'Heading 1'
        
        self.doc.add_paragraph()
        
        item_counter = 1
        for item in items:
            description = item.get('description', '')
            specifications = item.get('specifications', '')
            details = item.get('details', '')
            
            # Skip items with no technical content
            if not description and not specifications and not details:
                item_counter += 1
                continue
            
            # Item heading
            item_heading = self.doc.add_paragraph()
            run = item_heading.add_run(f"{item_counter}. {item.get('item_name', 'Item')}")
            run.font.size = Pt(12)
            run.font.bold = True
            run.font.name = self.font_name
            
            # Description
            if description:
                desc_para = self.doc.add_paragraph()
                run = desc_para.add_run(description)
                run.font.size = Pt(self.font_size_body)
                run.font.name = self.font_name
                self.doc.add_paragraph()
            
            # Specifications (as bullet points if structured)
            if specifications:
                spec_heading = self.doc.add_paragraph()
                run = spec_heading.add_run("Key Specifications:")
                run.font.bold = True
                run.font.size = Pt(self.font_size_body)
                run.font.name = self.font_name
                
                spec_para = self.doc.add_paragraph()
                run = spec_para.add_run(specifications)
                run.font.size = Pt(self.font_size_body)
                run.font.name = self.font_name
                self.doc.add_paragraph()
            
            # Additional details
            if details:
                details_para = self.doc.add_paragraph()
                run = details_para.add_run(details)
                run.font.size = Pt(self.font_size_small)
                run.font.name = self.font_name
                run.font.color.rgb = RGBColor(102, 102, 102)  # Gray text
                self.doc.add_paragraph()
            
            item_counter += 1
        
        self.doc.add_paragraph()
    
    def add_commercial_terms(self, company_data, supplier_terms=None):
        """
        Add commercial terms section
        Shows company standard terms + highlights differences from supplier
        """
        
        # Section heading
        heading = self.doc.add_paragraph()
        run = heading.add_run("Commercial Terms & Conditions")
        run.font.size = Pt(self.font_size_header)
        run.font.bold = True
        run.font.name = self.font_name
        heading.style = 'Heading 1'
        
        self.doc.add_paragraph()
        
        standard_terms = company_data.get('standard_terms', {})
        
        # Delivery Terms
        delivery_para = self.doc.add_paragraph()
        run = delivery_para.add_run("DELIVERY TERMS\n")
        run.font.bold = True
        run.font.size = Pt(self.font_size_body)
        run.font.name = self.font_name
        
        delivery_text = standard_terms.get('delivery', 'As per agreement')
        run = delivery_para.add_run(delivery_text)
        run.font.size = Pt(self.font_size_body)
        run.font.name = self.font_name
        
        # Note supplier difference if provided
        if supplier_terms and supplier_terms.get('delivery'):
            run = delivery_para.add_run(f"\nNote: Supplier delivery time is {supplier_terms['delivery']}")
            run.font.size = Pt(self.font_size_small)
            run.font.italic = True
            run.font.color.rgb = RGBColor(102, 102, 102)
        
        self.doc.add_paragraph()
        
        # Payment Terms
        payment_para = self.doc.add_paragraph()
        run = payment_para.add_run("PAYMENT TERMS\n")
        run.font.bold = True
        run.font.size = Pt(self.font_size_body)
        run.font.name = self.font_name
        
        payment_text = standard_terms.get('payment', 'As per agreement')
        run = payment_para.add_run(payment_text)
        run.font.size = Pt(self.font_size_body)
        run.font.name = self.font_name
        
        self.doc.add_paragraph()
        
        # Warranty
        warranty_para = self.doc.add_paragraph()
        run = warranty_para.add_run("WARRANTY\n")
        run.font.bold = True
        run.font.size = Pt(self.font_size_body)
        run.font.name = self.font_name
        
        warranty_text = standard_terms.get('warranty', 'As per manufacturer standard warranty')
        run = warranty_para.add_run(warranty_text)
        run.font.size = Pt(self.font_size_body)
        run.font.name = self.font_name
        
        self.doc.add_paragraph()
    
    def add_footer_section(self, company_data):
        """
        Add footer with bank details and legal information
        """
        
        # Separator line
        self.doc.add_paragraph("_" * 80)
        
        # Bank details
        bank_details = company_data.get('bank_details', {})
        
        if bank_details.get('bank_name') or bank_details.get('iban'):
            bank_para = self.doc.add_paragraph()
            run = bank_para.add_run("BANK DETAILS\n")
            run.font.bold = True
            run.font.size = Pt(self.font_size_small)
            run.font.name = self.font_name
            
            if bank_details.get('bank_name'):
                run = bank_para.add_run(f"Bank: {bank_details['bank_name']}\n")
                run.font.size = Pt(self.font_size_small)
                run.font.name = self.font_name
            
            if bank_details.get('iban'):
                run = bank_para.add_run(f"IBAN: {bank_details['iban']}\n")
                run.font.size = Pt(self.font_size_small)
                run.font.name = self.font_name
            
            if bank_details.get('swift'):
                run = bank_para.add_run(f"SWIFT: {bank_details['swift']}\n")
                run.font.size = Pt(self.font_size_small)
                run.font.name = self.font_name
            
            if bank_details.get('account_holder'):
                run = bank_para.add_run(f"Account Holder: {bank_details['account_holder']}\n")
                run.font.size = Pt(self.font_size_small)
                run.font.name = self.font_name
        
        # Legal info
        legal_para = self.doc.add_paragraph()
        run = legal_para.add_run("COMPANY REGISTRATION\n")
        run.font.bold = True
        run.font.size = Pt(self.font_size_small)
        run.font.name = self.font_name
        
        if company_data.get('registration_no'):
            run = legal_para.add_run(f"Registration No: {company_data['registration_no']}\n")
            run.font.size = Pt(self.font_size_small)
            run.font.name = self.font_name
        
        if company_data.get('tax_id'):
            run = legal_para.add_run(f"VAT No: {company_data['tax_id']}\n")
            run.font.size = Pt(self.font_size_small)
            run.font.name = self.font_name
        
        # Contact for questions
        if company_data.get('email') or company_data.get('phone'):
            contact_text = "For questions contact: "
            if company_data.get('email'):
                contact_text += company_data['email']
            if company_data.get('phone'):
                contact_text += f", {company_data['phone']}"
            
            contact_para = self.doc.add_paragraph()
            run = contact_para.add_run(contact_text)
            run.font.size = Pt(self.font_size_small)
            run.font.name = self.font_name
    
    def save(self, output_path):
        """Save the document"""
        self.doc.save(output_path)
        print(f"✓ Document saved: {output_path}", flush=True)
    
    # Helper methods
    
    def _set_cell_text(self, cell, text, bold=False, align="left", bg_color=None, text_color=None):
        """Set cell text with formatting"""
        cell.text = ""  # Clear existing
        para = cell.paragraphs[0]
        run = para.add_run(str(text))
        
        run.font.name = self.font_name
        run.font.size = Pt(self.font_size_body)
        run.font.bold = bold
        
        if text_color:
            if len(text_color) == 6:
                run.font.color.rgb = RGBColor(
                    int(text_color[0:2], 16),
                    int(text_color[2:4], 16),
                    int(text_color[4:6], 16)
                )
        
        if align == "center":
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "right":
            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
        if bg_color:
            self._set_cell_background(cell, bg_color)
    
    def _set_cell_background(self, cell, color_hex):
        """Set cell background color"""
        shading_elm = OxmlElement('w:shd')
        shading_elm.set(qn('w:fill'), color_hex)
        cell._element.get_or_add_tcPr().append(shading_elm)
    
    def _remove_table_borders(self, table):
        """Remove all borders from table"""
        tbl = table._element
        tblPr = tbl.tblPr
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)
        
        tblBorders = OxmlElement('w:tblBorders')
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'none')
            border.set(qn('w:sz'), '0')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), 'auto')
            tblBorders.append(border)
        
        tblPr.append(tblBorders)