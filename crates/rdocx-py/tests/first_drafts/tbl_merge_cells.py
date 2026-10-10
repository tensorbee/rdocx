# area: tables
# divergence: set_cell_grid_span
from rdocx import Document

doc = Document()
table = doc.add_table(rows=2, cols=3)
a = table.cell(0, 0)
b = table.cell(0, 2)
merged = a.merge(b)
merged.text = 'Merged header'
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert 'w:gridSpan w:val="3"' in xml and 'Merged header' in xml
