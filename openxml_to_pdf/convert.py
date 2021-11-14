import sys
import logging

import docx
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

from openxml_to_pdf import drawing

logger = logging.getLogger(__name__)
logger.addHandler(logging.StreamHandler(sys.stdin))
logger.setLevel(logging.INFO)

def iter_elements(parent):
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def convert(filename):
    doc = docx.Document(filename)
    pdf = drawing.init()

    pdf.set_author(doc.core_properties.author)
    
    for elem in iter_elements(doc):
        if isinstance(elem, Paragraph):
            drawing.draw_paragraph(doc, pdf, elem)
        elif isinstance(elem, Table):
            drawing.draw_table(doc, pdf, elem)
        else:
            raise ValueError("something's not right")

    pdf.output('output.pdf', 'F')

