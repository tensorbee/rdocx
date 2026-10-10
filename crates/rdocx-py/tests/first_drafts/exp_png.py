# area: export
# expected: pass
from rdocx import Document

doc = Document()
doc.add_paragraph('A page rendered to PNG.')
png = doc.render_page_to_png(0, dpi=72)
with open('page1.png', 'wb') as f:
    f.write(png)
# --- check
assert open('page1.png', 'rb').read(8) == b'\x89PNG\r\n\x1a\n'
