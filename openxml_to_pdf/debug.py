from docx.styles.style import _ParagraphStyle, _CharacterStyle
from docx.text.font import Font
from docx.text.parfmt import ParagraphFormat
from docx.text.tabstops import TabStops


def print_paragraph_style(style: _ParagraphStyle):
    print(f"""\nPARAGRAPH STYLE\n
name: {style.name}
type: {style.type}
hidden: {style.hidden}
locked: {style.locked}
priority: {style.priority}
base_style: {style.base_style}
style_id: {style.style_id}""")

def print_character_style(style: _CharacterStyle):
    print(f"""\nCHARACTER STYLE\n
name: {style.name}
type: {style.type}
hidden: {style.hidden}
locked: {style.locked}
priority: {style.priority}
style_id: {style.style_id}
base_style: {style.base_style}
""")

def print_paragraph_format(paragraph_format: ParagraphFormat):
    print(f"""\nPARAGRAPH FORMAT\n
alignment: {paragraph_format.alignment}
first_line_indent: {paragraph_format.first_line_indent}
keep_together: {paragraph_format.keep_together}
keep_with_next: {paragraph_format.keep_with_next}
left_indent: {paragraph_format.left_indent}
line_spacing: {paragraph_format.line_spacing}
line_spacing_rule: {paragraph_format.line_spacing_rule}
page_break_before: {paragraph_format.page_break_before}
right_indent: {paragraph_format.right_indent}
space_after: {paragraph_format.space_after}
space_before: {paragraph_format.space_before}
tab_stops: {print_tab_stops(paragraph_format.tab_stops)}
widow_control: {paragraph_format.widow_control}""")

def print_tab_stops(tab_stops: TabStops):
    return [f"{tab_stop.position}" for tab_stop in tab_stops]

def print_font(font: Font):
    print(f"""\nFONT\n
name: {font.name}
bold: {font.bold}
italic: {font.italic}
outline: {font.outline}
shadow: {font.shadow}
strike: {font.strike}
subscript: {font.subscript}
superscript: {font.superscript}
underline: {font.underline}
color: {font.color.rgb}
highlight_color: {font.highlight_color}
snap_to_grid: {font.snap_to_grid}
size: {font.size}
""")
