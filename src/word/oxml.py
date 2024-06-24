from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def insert_toc(paragraph):
    # Create the TOC field code
    fldSimple = OxmlElement('w:fldSimple')
    fldSimple.set(qn('w:instr'), 'TOC \\o "1-3" \\h \\z \\u')

    # Create the run element and append it to the paragraph
    run = OxmlElement('w:r')
    fldSimple.append(run)
    
    # Create the text element
    text = OxmlElement('w:t')
    text.text = "TOC will be inserted here"
    run.append(text)
    
    # Append the fldSimple element to the paragraph
    paragraph._element.append(fldSimple)

def copy_table(source_table, destination_document, csv_file):
    # Create a new table in the destination document
    destination_table = destination_document.add_table(rows=len(source_table.rows), cols=len(source_table.columns))
    
     # Open CSV file
    with open(csv_file, 'r', encoding='UTF-8') as file:
        csv_reader = csv.reader(file)
        for i, row in enumerate(csv_reader):
            if i < len(destination_table.rows):  # Ensure CSV rows don't exceed table rows
                for j, cell in enumerate(row):
                    if j < len(destination_table.columns):  # Ensure CSV columns don't exceed table columns
                        destination_table.cell(i, j).text = cell

    # Copy content and style from source table to destination table
    for i, row in enumerate(source_table.rows):
        for j, cell in enumerate(row.cells):
            # Copy cell style
            destination_table.cell(i, j)._element.get_or_add_tcPr().append(cell._element.tcPr)


