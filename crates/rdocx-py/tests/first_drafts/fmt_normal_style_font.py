# area: formatting
# divergence: set_style
from rdocx import Document
from rdocx.shared import Pt

doc = Document()
style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(11)
doc.add_paragraph('Body in Arial 11.')
doc.save('out.docx')
# --- check
styles = part('out.docx', 'word/styles.xml')
normal = styles[styles.index('w:styleId="Normal"'):]
normal = normal[:normal.index('</w:style>')]
assert 'w:ascii="Arial"' in normal and 'w:sz w:val="22"' in normal
