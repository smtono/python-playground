from oxml import OOXML
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE

# Define a function to add a hyperlink relationship
def add_hyperlink(paragraph, url, text, tooltip=None):
    part = paragraph.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = create_ooxml_element(OOXML.HYPERLINK, id=r_id)
    if tooltip:
        hyperlink.set(qn('w:tooltip'), tooltip)
    
    run = create_ooxml_element(OOXML.RUN)
    r_pr = create_ooxml_element(OOXML.STYLE_RUN_PROPERTIES)
    r_style = create_ooxml_element(OOXML.STYLE, val='Hyperlink')
    r_pr.append(r_style)
    run.append(r_pr)

    text_element = create_ooxml_element(OOXML.TEXT)
    text_element.text = text
    run.append(text_element)
    hyperlink.append(run)
    paragraph._element.append(hyperlink)