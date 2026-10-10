# area: links
# needs: #322
from rdocx import Document

doc = Document()
doc.add_paragraph('Intro')
p = doc.add_paragraph('Jump to ')
heading = doc.add_heading('The end', level=1)
p.add_hyperlink('the end', anchor=heading)
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert re.search(r'w:anchor="([^"]+)"', xml) and '<w:bookmarkStart' in xml
