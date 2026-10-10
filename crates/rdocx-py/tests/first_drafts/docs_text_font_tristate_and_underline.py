# area: docs
# expected: pass
# python-docx text.rst, tri-state font properties and underline
from rdocx import Document
from rdocx.enum.text import WD_UNDERLINE

document = Document()
font = document.add_paragraph('').add_run('styled').font
font.italic = True
font.italic = False
font.italic = None
font.underline = True
font.underline = WD_UNDERLINE.DOT_DASH
document.save('out.docx')
# --- check
xml = part('out.docx')
assert 'w:u w:val="dotDash"' in xml and '<w:i' not in xml
