from typing import Optional, List
from fpdf import FPDF

from docx.document import Document
from docx.text.font import Font

from openxml_to_pdf import helpers

def set_font(pdf: FPDF, family: str, style: str, size: int):
    try:
        pdf.set_font(family, style, size)
    except RuntimeError:
        helpers.add_font(pdf, family, style)
        pdf.set_font(family, style, size)

def get_color(font: Font):
    rgb = font.color.rgb
    if rgb is None:
        return None
    return {
        'r': rgb[0],
        'g': rgb[1],
        'b': rgb[2],
    }

def get_fill(font: Font) -> Optional[dict]:
    color = font.highlight_color
    if color is None:
        return None
    return helpers.convert_wd_color_index(color)

def get_style(font: Font) -> Optional[str]:
    style = ''
    if font.bold:
        style += 'b'
    if font.italic:
        style += 'i'
    if font.underline:
        style += 'u'
    return style if style else None

def get_size(font: Font):
    if font.size is None:
        return None
    return font.size.pt

def get_family(font: Font):
    if font is None:
        return None
    return font.name

def get_all_caps(*args: List[Font]):
    fonts = list(args[::-1])
    while fonts:
        font = fonts.pop()
        if font.all_caps is not None:
            return font.all_caps
    return False
