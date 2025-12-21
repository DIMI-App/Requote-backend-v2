"""
Standard Offer 3 Template Structure
Defines the professional quotation format that will be generated
"""

import os
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
        
    def copy_header_footer_from_template(self, template_path):
        """
        Copy header and footer directly from Offer 2 template to Offer 3
        This preserves logos, formatting, company info exactly as-is
        """
        
        print("\n" + "=" * 60, flush=True)
        print("COPYING HEADER AND FOOTER FROM TEMPLATE", flush=True)
        print("=" * 60, flush=True)
        
        if not os.path.exists(template_path):
            print(f"⚠ Template not found: {template_path}", flush=True)
            return False
        
        try:
            from docx import Document as DocxDocument
            template_doc = DocxDocument(template_path)
            
            # Copy header from each section
            for section_idx, template_section in enumerate(template_doc.sections):
                # Get or create corresponding section in our document
                if section_idx >= len(self.doc.sections):
                    self.doc.add_section()
                
                our_section = self.doc.sections[section_idx]
                
                # Copy header
                if template_section.header:
                    print(f"Copying header from section {section_idx + 1}...", flush=True)
                    
                    # Clear existing header
                    our_section.header._element.clear_content()
                    
                    # Copy all header content
                    for element in template_section.header._element:
                        our_section.header._element.append(element)
                    
                    print(f"✓ Header copied from section {section_idx + 1}", flush=True)
                
                # Copy footer
                if template_section.footer:
                    print(f"Copying footer from section {section_idx + 1}...", flush=True)
                    
                    # Clear existing footer
                    our_section.footer._element.clear_content()
                    
                    # Copy all footer content
                    for element in template_section.footer._element:
                        our_section.footer._element.append(element)
                    
                    print(f"✓ Footer copied from section {section_idx + 1}", flush=True)
            
            print("=" * 60, flush=True)
            return True
            
        except Exception as e:
            print(f"✗ Error copying header/footer: {str(e)}", flush=True)
            import traceback
            traceback.print_exc()
            return False
    
    def add_header_section(self, company_data):
        """
        DEPRECATED - This method is no longer used
        We now copy header/footer directly from template
        Keeping this for backwards compatibility
        """
        print("⚠ add_header_section called but header is copied from template", flush=True)
        pass
    
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
        DEPRECATED - Footer is now copied directly from template
        Keeping this for backwards compatibility
        """
        print("⚠ add_footer_section called but footer is copied from template", flush=True)
        pass
    
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