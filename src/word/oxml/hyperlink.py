from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml.shared import OxmlElement

# Define a function to add a hyperlink relationship
def add_hyperlink(paragraph, url, text, tooltip=None):
    part = paragraph.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = create_ooxml_element(OOXMLTags.HYPERLINK, id=r_id)
    if tooltip:
        hyperlink.set(qn('w:tooltip'), tooltip)
    
    run = create_ooxml_element(OOXMLTags.RUN)
    r_pr = create_ooxml_element(OOXMLTags.STYLE_RUN_PROPERTIES)
    r_style = create_ooxml_element(OOXMLTags.STYLE, val='Hyperlink')
    r_pr.append(r_style)
    run.append(r_pr)

    text_element = create_ooxml_element(OOXMLTags.TEXT)
    text_element.text = text
    run.append(text_element)
    hyperlink.append(run)
    paragraph._element.append(hyperlink)

# Add a references page with hyperlinks
doc.add_heading('References', level=1)
p = doc.add_paragraph('For more information, visit ')
add_hyperlink(p, 'https://www.openai.com', 'OpenAI')

# Save the document
doc.save('complete_document_with_hyperlinks.docx')
