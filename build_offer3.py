def add_structured_content_to_doc(doc, items):
    """
    Add technical content with PRESERVED STRUCTURE
    """
    
    heading = doc.add_paragraph()
    run = heading.add_run("Technical Specifications")
    run.font.size = Pt(14)
    run.font.bold = True
    
    doc.add_paragraph()
    
    for item in items:
        content_blocks = item.get('content_blocks', [])
        
        if not content_blocks:
            continue
        
        # Item heading
        item_heading = doc.add_paragraph()
        run = item_heading.add_run(f"{item['item_number']}. {item['item_name']}")
        run.font.size = Pt(12)
        run.font.bold = True
        
        # Rebuild content with structure
        for block in content_blocks:
            block_type = block['type']
            
            if block_type == 'table':
                # Rebuild table
                table_data = block['data']
                num_rows = len(table_data)
                num_cols = len(table_data[0]) if table_data else 0
                
                if num_rows > 0 and num_cols > 0:
                    table = doc.add_table(rows=num_rows, cols=num_cols)
                    table.style = 'Light Grid Accent 1'
                    
                    for row_idx, row_data in enumerate(table_data):
                        for col_idx, cell_text in enumerate(row_data):
                            if col_idx < num_cols:
                                table.rows[row_idx].cells[col_idx].text = str(cell_text or '')
            
            elif block_type == 'bullet':
                # Bullet point
                para = doc.add_paragraph(style='List Bullet')
                run = para.add_run(block['text'])
                run.font.size = Pt(11)
            
            elif block_type == 'numbered_list':
                # Numbered list
                para = doc.add_paragraph(style='List Number')
                run = para.add_run(block['text'])
                run.font.size = Pt(11)
            
            elif block_type == 'heading':
                # Sub-heading
                para = doc.add_paragraph()
                run = para.add_run(block['text'])
                run.font.size = Pt(11)
                run.font.bold = True
            
            else:
                # Normal paragraph
                para = doc.add_paragraph()
                run = para.add_run(block['text'])
                run.font.size = Pt(11)
        
        doc.add_paragraph()  # Spacing between items