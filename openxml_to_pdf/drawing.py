from typing import List
import logging

from fpdf import FPDF

from docx.document import Document
from docx.table import Table, _Row, _Cell
from docx.text.paragraph import Paragraph, Run
from docx.opc.coreprops import CoreProperties
from docx.section import Section

from openxml_to_pdf import styles
from openxml_to_pdf import debug
from openxml_to_pdf import helpers

logger = logging.getLogger(__name__)
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)

WIDTH = 210
HEIGHT = 297
INCH = 25.4

def init(doc: Document):
    def get_page_format(doc: Document):
        section = doc.sections[0]
        sizes = [section.page_width.mm, section.page_height.mm]
        return (min(sizes), max(sizes))

    def set_metadata(pdf: FPDF, properties: CoreProperties):
        pdf.set_author(properties.author)

    def set_margin(pdf: FPDF, section: Section):
        pdf.set_top_margin(section.top_margin.pt // 2)
        pdf.set_right_margin(section.right_margin.pt // 2)
        pdf.set_left_margin(section.left_margin.pt // 2)

    page_format = get_page_format(doc)
    pdf = FPDF(orientation='P', unit='mm', format=page_format)

    set_margin(pdf, doc.sections[0])
    set_metadata(pdf, doc.core_properties)

    # Set compression to False, to prevent latin-1 encoding.
    pdf.set_compression(False)

    # Add initial page.
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
    default_run_style = doc.styles.default(2)
    default_paragraph_style = doc.styles.default(1)
    def get_color(doc: Document, paragraph: Paragraph, run: Run):
        return styles.get_color(run.font) or \
            styles.get_color(run.style.font) or \
            styles.get_color(default_run_style.font) or \
            styles.get_color(paragraph.style.font) or \
            styles.get_color(default_paragraph_style.font) or \
            {'r': 0, 'g': 0, 'b': 0}

    def get_fill(doc: Document, paragraph: Paragraph, run: Run):
        return styles.get_fill(run.font) or \
            styles.get_fill(run.style.font) or \
            styles.get_fill(default_run_style.font) or \
            styles.get_fill(paragraph.style.font) or \
            styles.get_fill(default_paragraph_style.font)

    def get_style(doc: Document, paragraph: Paragraph, run: Run):
        return styles.get_style(run.font) or \
            styles.get_style(run.style.font) or \
            styles.get_style(default_run_style.font) or \
            styles.get_style(paragraph.style.font) or \
            styles.get_style(default_paragraph_style.font) or \
            ''

    def get_size(doc: Document, paragraph: Paragraph, run: Run):
        return styles.get_size(run.font) or \
            styles.get_size(run.style.font) or \
            styles.get_size(default_run_style.font) or \
            styles.get_size(paragraph.style.font) or \
            styles.get_size(default_paragraph_style.font) or \
            10

    def get_family(doc: Document, paragraph: Paragraph, run: Run):
        return styles.get_family(run.font) or \
            styles.get_family(run.style.font) or \
            styles.get_family(default_run_style.font) or \
            styles.get_family(paragraph.style.font) or \
            styles.get_family(default_paragraph_style.font) or \
            'Arial'

    def get_all_caps(doc: Document, paragraph: Paragraph, run: Run):
        return styles.get_all_caps(
            run.font,
            run.style.font,
            default_run_style.font,
            paragraph.style.font,
            default_paragraph_style.font,
        )

    if paragraph.paragraph_format.page_break_before:
        pdf.add_page()

    if not paragraph.runs:
        return

    paragraph_style = paragraph.style
    for run in paragraph.runs:
        run_style = run.style

        color = get_color(doc, paragraph, run)
        fill = get_fill(doc, paragraph, run)
        style = get_style(doc, paragraph, run)
        size = get_size(doc, paragraph, run)
        family = get_family(doc, paragraph, run)
        all_caps = get_all_caps(doc, paragraph, run)

        pdf.set_text_color(**color)
        if fill:
            pdf.set_fill_color(**fill)

        styles.set_font(pdf, family, style, size)
        text = run.text if not all_caps else run.text.upper()

        # TOOD: get line height from document.
        pdf.write(6, text)
    pdf.ln()
