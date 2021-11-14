from typing import Optional
from fpdf import FPDF

from docx.document import Document
from docx.text.font import Font
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX

from openxml_to_pdf import helpers

def apply_font(doc: Document, pdf: FPDF, text: str, font: Font) -> str:
    if font.all_caps:
        text = text.upper()

    default = doc.styles.default(2)
    
    color = _get_color(default.font.color.rgb) or _get_color(font.color.rgb) or {'r': 0, 'g': 0, 'b': 0}
    fill = _get_fill(default.font.highlight_color) or _get_fill(font.highlight_color)
    style = _get_style(default.font) or _get_style(font) or ''

    size = (default.font.size.pt if default.font.size else None) or (font.size.pt if font.size else None) or 8
    family = (default.font.name or font.name or 'Arial').lower()

    try:
        pdf.set_font(family, style, size)
    except RuntimeError:
        helpers.add_font(pdf, family, style)
        pdf.set_font(family, style, size)

    pdf.set_text_color(**color)
    if fill:
        pdf.set_fill_color(**fill)
    return text

def _get_color(rgb: Optional[RGBColor]):
    if rgb is None:
        return None
    return {
        'r': rgb[0],
        'g': rgb[1],
        'b': rgb[2],
    }

def _get_fill(color: Optional[WD_COLOR_INDEX]) -> Optional[dict]:
    if color is None:
        return None
    return helpers.convert_wd_color_index(color)

def _get_style(font: Font) -> Optional[str]:
    style = ''
    if font.bold:
        style += 'b'
    if font.italic:
        style += 'i'
    if font.underline:
        style += 'u'
    return style if style else None

