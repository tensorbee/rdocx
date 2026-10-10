# area: document
# expected: pass
import io
from rdocx import Document

with open('template.docx', 'rb') as f:
    doc = Document(f)
doc.add_paragraph('From a stream')
buffer = io.BytesIO()
doc.save(buffer)
with open('out.docx', 'wb') as f:
    f.write(buffer.getvalue())
# --- check
assert 'From a stream' in part('out.docx')
