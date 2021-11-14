import sys
import logging

from fpdf import FPDF

import docx
from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.opc.coreprops import CoreProperties
from docx.section import Section

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
    pdf = drawing.init(doc)
    _set_metadata(pdf, doc.core_properties)
    _set_margin(pdf, doc.sections[0])
    
    for elem in iter_elements(doc):
        if isinstance(elem, Paragraph):
            drawing.draw_paragraph(doc, pdf, elem)
        elif isinstance(elem, Table):
            drawing.draw_table(doc, pdf, elem)
        else:
            raise ValueError("something's not right")

    pdf.output('output.pdf', 'F')

def _set_metadata(pdf: FPDF, properties: CoreProperties):
    pdf.set_author(properties.author)

def _set_margin(pdf: FPDF, section: Section):
    pdf.set_top_margin(section.top_margin.pt // 2)
    pdf.set_right_margin(section.right_margin.pt // 2)
    pdf.set_left_margin(section.left_margin.pt // 2)
