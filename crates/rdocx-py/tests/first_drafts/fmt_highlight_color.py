# area: formatting
# expected: pass
from rdocx import Document
from rdocx.enum.text import WD_COLOR_INDEX

doc = Document()
run = doc.add_paragraph('').add_run('look here')
run.font.highlight_color = WD_COLOR_INDEX.YELLOW
doc.save('out.docx')
# --- check
assert 'w:highlight w:val="yellow"' in part('out.docx')
