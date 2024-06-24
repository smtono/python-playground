from enum import Enum
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

class OOXMLTags(Enum):
    # Document properties
    TITLE = "w:title"
    SUBJECT = "w:subject"
    AUTHOR = "w:author"
    DESCRIPTION = "w:description"
    KEYWORDS = "w:keywords"
    CATEGORY = "w:category"
    CONTENT_STATUS = "w:contentStatus"
    LAST_MODIFIED_BY = "w:lastModifiedBy"
    REVISION = "w:revision"
    LAST_PRINTED = "w:lastPrinted"
    CREATED = "w:created"
    MODIFIED = "w:modified"

    # Body elements
    PARAGRAPH = "w:p"
    RUN = "w:r"
    TEXT = "w:t"
    BREAK = "w:br"
    TAB = "w:tab"
    BOOKMARK_START = "w:bookmarkStart"
    BOOKMARK_END = "w:bookmarkEnd"
    COMMENT_RANGE_START = "w:commentRangeStart"
    COMMENT_RANGE_END = "w:commentRangeEnd"
    COMMENT_REFERENCE = "w:commentReference"
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
    PAGE_SIZE = "w:pgSz"
    PAGE_MARGINS = "w:pgMar"
    COLUMNS = "w:cols"

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
    element = OxmlElement(tag.value)
    for key, value in attributes.items():
        element.set(qn(key), value)
    return element

def create_custom_field(tag):
    element = create_ooxml_element(OOXMLTags.FIELD, instr=tag.value)
    return element
