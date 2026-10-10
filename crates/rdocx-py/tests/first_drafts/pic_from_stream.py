# area: pictures
# needs: #322
import io
from rdocx import Document
from rdocx.shared import Inches

with open('logo.png', 'rb') as f:
    stream = io.BytesIO(f.read())
doc = Document()
doc.add_picture(stream, height=Inches(0.5))
doc.save('out.docx')
# --- check
assert 'cy="457200"' in part('out.docx')
