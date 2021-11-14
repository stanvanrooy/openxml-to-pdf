from typing import List
import logging

from fpdf import FPDF

from docx.document import Document
from docx.table import Table, _Row, _Cell
from docx.text.paragraph import Paragraph

from openxml_to_pdf import styles
from openxml_to_pdf import debug

logger = logging.getLogger(__name__)
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.INFO)

WIDTH = 210
HEIGHT = 297
INCH = 25.4

def init(doc: Document):
    page_format = _get_page_format(doc)
    pdf = FPDF(orientation='P', unit='mm', format=page_format)
    pdf.add_page()
    return pdf

def draw_table(doc: Document, pdf: FPDF, table: Table):
    for row in table.rows:
        draw_row(pdf, row)

def draw_row(pdf: FPDF, row: _Row):
    for cell in row.cells:
        draw_cell(pdf, cell)

def draw_cell(pdf: FPDF, cell: _Cell):
    return
    # print(cell)

def draw_paragraph(doc: Document, pdf: FPDF, paragraph: Paragraph):
    # If this paragraph is after a page break, add a new page.
    if paragraph.paragraph_format.page_break_before:
        pdf.add_page()

    if not paragraph.runs:
        return

    default_style = doc.styles.default(1)
    for run in paragraph.runs:
        text = styles.apply_font(doc, pdf, run.text, run.font)
        pdf.write(paragraph.paragraph_format.line_spacing or 1, text)
    pdf.ln()

def _get_page_format(doc: Document):
    section = doc.sections[0]
    sizes = [section.page_width.mm, section.page_height.mm]
    return (min(sizes), max(sizes))
