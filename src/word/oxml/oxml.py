"""
prepend calls with "add_paragraph" for the tag to have a container in the document
# Creating standard OOXML elements
paragraph = create_ooxml_element(OOXMLTags.PARAGRAPH, align='center')
run = create_ooxml_element(OOXMLTags.RUN)
text = create_ooxml_element(OOXMLTags.TEXT)

# Creating custom instructions using easier names
page_field = create_ooxml_element(OOXMLCustomInstructions.PAGE)
toc_field = create_ooxml_element(OOXMLCustomInstructions.TABLE_OF_CONTENTS)
"""

from enum import Enum
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

class OOXMLTags(Enum):
    # Document properties
    TITLE = "w:title"
    SUBJECT = "w:subject"
    AUTHOR = "w:author"
    DESCRIPTION = "w:description"

    # Body elements
    PARAGRAPH = "w:p"
    RUN = "w:r"
    TEXT = "w:t"
    BREAK = "w:br"
    TAB = "w:tab"
    BOOKMARK_START = "w:bookmarkStart"
    BOOKMARK_END = "w:bookmarkEnd"
    FOOTNOTE_REFERENCE = "w:footnoteReference"
    ENDNOTE_REFERENCE = "w:endnoteReference"
    FIELD = "w:fldSimple"
    HYPERLINK = "w:hyperlink"

    # Table elements
    TABLE = "w:tbl"
    TABLE_ROW = "w:tr"
    TABLE_CELL = "w:tc"
    TABLE_GRID = "w:tblGrid"
    TABLE_HEADER = "w:tblHeader"
    TABLE_FOOTER = "w:tblFooter"
    TABLE_PROPERTIES = "w:tblPr"
    TABLE_ROW_PROPERTIES = "w:trPr"
    TABLE_CELL_PROPERTIES = "w:tcPr"
    TABLE_STYLE = "w:tblStyle"

    # Styles
    STYLE = "w:style"
    STYLE_NAME = "w:name"
    BASE_STYLE = "w:basedOn"
    NEXT_STYLE = "w:next"
    LINKED_STYLE = "w:link"
    STYLE_PARAGRAPH_PROPERTIES = "w:pPr"
    STYLE_RUN_PROPERTIES = "w:rPr"
    STYLE_TABLE_PROPERTIES = "w:tblPr"
    STYLE_TABLE_ROW_PROPERTIES = "w:trPr"
    STYLE_TABLE_CELL_PROPERTIES = "w:tcPr"
    STYLE_TYPE = "w:type"

    # Sections
    SECTION_PROPERTIES = "w:sectPr"
    HEADER_REFERENCE = "w:headerReference"
    FOOTER_REFERENCE = "w:footerReference"

    # Lists
    NUMBERING = "w:numbering"
    NUMBERING_PROPERTIES = "w:numPr"
    NUMBERING_ID = "w:numId"
    ABSTRACT_NUM_ID = "w:abstractNumId"
    LEVEL = "w:lvl"
    LEVEL_TEXT = "w:lvlText"
    LEVEL_JC = "w:lvlJc"
    LEVEL_PPR = "w:lvlPPr"
    LEVEL_RPR = "w:lvlRPr"
    BULLET = "w:bullets"
    NUMBERING_STYLE = "w:numStyleLink"

    # Misc
    DRAWING = "w:drawing"
    INLINE = "w:inline"
    ANCHOR = "w:anchor"
    EXTENT = "w:extent"
    EFFECT_EXTENT = "w:effectExtent"
    WRAP_NONE = "w:wrapNone"
    WRAP_SQUARE = "w:wrapSquare"
    WRAP_TIGHT = "w:wrapTight"
    WRAP_THROUGH = "w:wrapThrough"
    WRAP_TOP_AND_BOTTOM = "w:wrapTopAndBottom"
    PICTURE = "w:pic"
    SHAPE = "w:sp"

class OOXMLCustomInstructions(Enum):
    TABLE_OF_CONTENTS = r'TOC \o "1-3" \h \z \u'
    TABLE_OF_FIGURES = r'TOC \c "Figure" \h \z \u'
    TABLE_OF_TABLES = r'TOC \c "Table" \h \z \u'
    PAGE = 'PAGE'
    NUMPAGES = 'NUMPAGES'
    DATE = 'DATE'
    TIME = 'TIME'
    FILENAME = 'FILENAME'
    AUTHOR = 'AUTHOR'
    DOCPROPERTY = 'DOCPROPERTY'
    REF = 'REF'
    NOTEREF = 'NOTEREF'

def create_ooxml_element(tag, **attributes):
    if isinstance(tag, OOXMLTags):
        element = OxmlElement(tag.value)
        for key, value in attributes.items():
            element.set(qn(key), value)
        return element
    elif isinstance(tag, OOXMLCustomInstructions):
        element = OxmlElement(OOXMLTags.FIELD.value)
        element.set(qn('instr'), tag.value)
        return element
    else:
        raise ValueError(f"Unsupported tag: {tag}")
