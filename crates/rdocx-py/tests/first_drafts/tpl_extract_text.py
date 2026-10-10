# area: templates
# expected: pass
from rdocx import Document

doc = Document('template.docx')
text = '\n'.join(p.text for p in doc.paragraphs)
cells = [cell.text for table in doc.tables for row in table.rows for cell in row.cells]
with open('out.txt', 'w') as f:
    f.write(text + '\n' + '\n'.join(cells))
# --- check
content = open('out.txt').read()
assert 'Dear {{name}},' in content and 'Metric' in content
