from enum import Enum
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

class OOXMLTag(Enum):
    # Document properties
    TITLE = "title"
    SUBJECT = "subject"
    AUTHOR = "author"
    DESCRIPTION = "description"
    KEYWORDS = "keywords"
    CATEGORY = "category"
    CONTENT_STATUS = "contentStatus"
    LAST_MODIFIED_BY = "lastModifiedBy"
    REVISION = "revision"
    LAST_PRINTED = "lastPrinted"
    CREATED = "created"
    MODIFIED = "modified"

    # Body elements
    PARAGRAPH = "p"
    RUN = "r"
    TEXT = "t"
    BREAK = "br"
    TAB = "tab"
    BOOKMARK_START = "bookmarkStart"
    BOOKMARK_END = "bookmarkEnd"
    COMMENT_RANGE_START = "commentRangeStart"
    COMMENT_RANGE_END = "commentRangeEnd"
    COMMENT_REFERENCE = "commentReference"
    FOOTNOTE_REFERENCE = "footnoteReference"
    ENDNOTE_REFERENCE = "endnoteReference"
    FIELD = "fldSimple"
    HYPERLINK = "hyperlink"

    # Table elements
    TABLE = "tbl"
    TABLE_ROW = "tr"
    TABLE_CELL = "tc"
    TABLE_GRID = "tblGrid"
    TABLE_HEADER = "tblHeader"
    TABLE_FOOTER = "tblFooter"
    TABLE_PROPERTIES = "tblPr"
    TABLE_ROW_PROPERTIES = "trPr"
    TABLE_CELL_PROPERTIES = "tcPr"
    TABLE_STYLE = "tblStyle"

    # Styles
    STYLE = "style"
    STYLE_NAME = "name"
    BASE_STYLE = "basedOn"
    NEXT_STYLE = "next"
    LINKED_STYLE = "link"
    STYLE_PARAGRAPH_PROPERTIES = "pPr"
    STYLE_RUN_PROPERTIES = "rPr"
    STYLE_TABLE_PROPERTIES = "tblPr"
    STYLE_TABLE_ROW_PROPERTIES = "trPr"
    STYLE_TABLE_CELL_PROPERTIES = "tcPr"
    STYLE_TYPE = "type"

    # Sections
    SECTION_PROPERTIES = "sectPr"
    HEADER_REFERENCE = "headerReference"
    FOOTER_REFERENCE = "footerReference"
    PAGE_SIZE = "pgSz"
    PAGE_MARGINS = "pgMar"
    COLUMNS = "cols"

    # Lists
    NUMBERING = "numbering"
    NUMBERING_PROPERTIES = "numPr"
    NUMBERING_ID = "numId"
    ABSTRACT_NUM_ID = "abstractNumId"
    LEVEL = "lvl"
    LEVEL_TEXT = "lvlText"
    LEVEL_JC = "lvlJc"
    LEVEL_PPR = "lvlPPr"
    LEVEL_RPR = "lvlRPr"
    BULLET = "bullets"
    NUMBERING_STYLE = "numStyleLink"

    # Misc
    DRAWING = "drawing"
    INLINE = "inline"
    ANCHOR = "anchor"
    EXTENT = "extent"
    EFFECT_EXTENT = "effectExtent"
    WRAP_NONE = "wrapNone"
    WRAP_SQUARE = "wrapSquare"
    WRAP_TIGHT = "wrapTight"
    WRAP_THROUGH = "wrapThrough"
    WRAP_TOP_AND_BOTTOM = "wrapTopAndBottom"
    PICTURE = "pic"
    SHAPE = "sp"

class OOXMLInstruction(Enum):
    TABLE_OF_CONTENTS = r'TOC \o "1-3" \h \z \u'
    TABLE_OF_FIGURES = r'TOC \c "Figure" \h \z \u'
    TABLE_OF_TABLES = r'TOC \c "Table" \h \z \u'

def create_ooxml_element(tag: OOXMLTag, **attributes):
    element = OxmlElement(f'w:{tag.value}')
    for key, value in attributes.items():
        element.set(qn(f'w:{key}'), value)
    return element

def create_custom_field(tag: OOXMLInstruction):
    element = create_ooxml_element(OOXMLTag.FIELD, instr=tag.value)
    return element
