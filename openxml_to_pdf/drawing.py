from typing import List
import logging

from fpdf import FPDF

from docx.table import Table, _Row, _Cell
from docx.text.paragraph import Paragraph

from openxml_to_pdf import styles

logger = logging.getLogger(__name__)
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.INFO)

WIDTH = 210
HEIGHT = 297
INCH = 25.4

def init():
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font('Times', '', 10.0)
    return pdf

def draw_table(pdf: FPDF, table: Table):
    for row in table.rows:
        draw_row(pdf, row)

def draw_row(pdf: FPDF, row: _Row):
    for cell in row.cells:
        draw_cell(pdf, cell)

def draw_cell(pdf: FPDF, cell: _Cell):
    return
    # print(cell)

def draw_paragraph(pdf: FPDF, paragraph: Paragraph):
    # If this paragraph is after a page break, add a new page.
    if paragraph.paragraph_format.page_break_before:
        pdf.add_page()

    for run in paragraph.runs:
        text = styles.apply_font(pdf, run.text, run.font)
        pdf.write(10, text)

