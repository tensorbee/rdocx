# area: headers-footers
# needs: #304
from rdocx import Document

doc = Document()
section = doc.sections[0]
section.different_first_page_header_footer = True
section.first_page_header.paragraphs[0].text = 'Cover'
section.header.paragraphs[0].text = 'Running head'
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert '<w:titlePg/>' in xml and 'w:type="first"' in xml
