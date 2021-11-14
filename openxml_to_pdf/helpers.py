from typing import Optional
import os
import re
import subprocess
import sys

from docx.enum.text import WD_COLOR_INDEX
from fpdf import FPDF

WD_COLOR_INDEX_MAP = {
    WD_COLOR_INDEX.BLACK: '#000000',
    WD_COLOR_INDEX.BLUE: '#0000FF',
    WD_COLOR_INDEX.BRIGHT_GREEN: '#00FF00',
    WD_COLOR_INDEX.DARK_BLUE: '#000080',
    WD_COLOR_INDEX.DARK_RED: '#800000',
    WD_COLOR_INDEX.DARK_YELLOW: '#808000',
    WD_COLOR_INDEX.GRAY_25: '#808080',
    WD_COLOR_INDEX.GRAY_50: '#f7f7f7',
    WD_COLOR_INDEX.GREEN: '#008000',
    WD_COLOR_INDEX.PINK: '#ffc0cb',
    WD_COLOR_INDEX.RED: '#ff0000',
    WD_COLOR_INDEX.TURQUOISE: '#40e0d0',
    WD_COLOR_INDEX.VIOLET: '#ee82ee',
    WD_COLOR_INDEX.WHITE: '#ffffff',
    WD_COLOR_INDEX.YELLOW: '#FFFF00',
    WD_COLOR_INDEX.TEAL: '#008080',
}


def convert_wd_color_index(color: WD_COLOR_INDEX) -> dict:
    """Convert a Word color index to a hex color string."""
    hex = WD_COLOR_INDEX_MAP[color]
    return _convert_hex_to_rgb(hex)

def add_font(pdf: FPDF, family: str, style: str):
    """Add a font to the PDF."""
    path = _get_font_path(family, style)
    pdf.add_font(family.lower(), style, path, uni=True)

def _get_font_path(family: str, style: str) -> Optional[str]:
    """Use fc-match to get the absolute path to the font family"""
    family_with_style = _add_style_to_family(family, style)
    try:
        font_path = subprocess.check_output(
            ['fc-match', '-f', '%{file}', family_with_style],
            stderr=subprocess.STDOUT,
        ).decode('utf-8').strip()
    except subprocess.CalledProcessError as e:
        print(e.output)
        sys.exit(1)

    if font_path is None:
        return None
    return font_path

def _convert_hex_to_rgb(hex: str) -> dict:
    """Convert a hex color string to an RGB color dict."""
    hex = hex.lstrip('#')
    return {
        'r': int(hex[0:2], 16),
        'g': int(hex[2:4], 16),
        'b': int(hex[4:6], 16),
    }

def _add_style_to_family(family: str, style: str) -> str:
    """Add a style to a font family."""
    style = style.lower()
    if 'b' in style:
        family += ':bold'
    if 'i' in style:
        family += ':italic'
    if 'u' in style:
        family += ':underline'
    return family
