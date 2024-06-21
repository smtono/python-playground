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

# Load the document
doc_path = 'populated_document.docx'
doc = Document(doc_path)

# Insert a paragraph for the TOC at the beginning of the document
toc_paragraph = doc.paragraphs[0].insert_paragraph_before()
insert_toc(toc_paragraph)

# Save the document
new_doc_path = 'document_with_toc.docx'
doc.save(new_doc_path)
