# area: export
# expected: pass
from rdocx import Document

doc = Document()
doc.add_heading('Report', level=1)
doc.add_paragraph('Exported to PDF.')
pdf = doc.to_pdf()
with open('out.pdf', 'wb') as f:
    f.write(pdf)
# --- check
assert open('out.pdf', 'rb').read(5) == b'%PDF-'
