from docx import Document
from oxml import OOXMLTag, OOXMLInstruction
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE
from oxml import create_custom_field
# Add a Table of Contents
def add_table_of_contents(doc: Document):
    doc.add_paragraph('Table of Contents', style='Heading 1')
    toc = create_custom_field(OOXMLInstruction.TABLE_OF_CONTENTS)
    p = doc.add_paragraph()
    p._element.append(toc)


