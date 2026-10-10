# area: headers-footers
# expected: pass
# rdocx's one-call header and footer, as an agent finds them in the rdocx docs
from rdocx import Document

doc = Document()
doc.set_header('Draft')
doc.set_footer('Page footer')
doc.save('out.docx')
# --- check
hf = [part('out.docx', n) for n in names('out.docx') if n.startswith(('word/header', 'word/footer'))]
assert any('Draft' in x for x in hf) and any('Page footer' in x for x in hf)
