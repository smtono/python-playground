from oxml import OOXML, create_ooxml_element
from docx import Document

from oxml_util import add_hyperlink

# Create a new Document
doc = Document()

# Add a title
doc.add_heading('Document Title', level=1)

# Add a Table of Contents
doc.add_paragraph('Table of Contents', style='Heading 1')
toc = create_ooxml_element(OOXML.FIELD, instr=r'TOC \o "1-3" \h \z \u')
p = doc.add_paragraph()
p._element.append(toc)

# Add a Table of Figures
doc.add_paragraph('Table of Figures', style='Heading 1')
tof = create_ooxml_element(OOXML.FIELD, instr=r'TOC \c "Figure" \h \z \u')
p = doc.add_paragraph()
p._element.append(tof)

# # Add a Table of Tables
# doc.add_paragraph('Table of Tables', style='Heading 1')
# tot = create_ooxml_element(OOXML.FIELD, instr=r'TOC \c "Table" \h \z \u')
# p = doc.add_paragraph()
# p._element.append(tot)

# Add sections with multiple paragraphs
for i in range(1, 4):
    doc.add_heading(f'Section {i}', level=1)
    for j in range(1, 4):
        doc.add_heading(f'Subsection {i}.{j}', level=2)
        for k in range(1, 4):
            doc.add_paragraph(f'This is paragraph {i}.{j}.{k}.')

# # Add a references page with hyperlinks
# doc.add_heading('References', level=1)
# p = doc.add_paragraph('For more information, visit ')
# add_hyperlink(p, 'https://www.openai.com', 'OpenAI')

# Save the document
doc.save('complete_document.docx')
