# area: shapes
# expected: pass
from rpptx import Presentation
from rpptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
tf = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1)).text_frame
run = tf.paragraphs[0].add_run()
run.text = 'Visit us'
run.hyperlink.address = 'https://example.com'
prs.save('out.pptx')
# --- check
assert '<a:hlinkClick' in part('out.pptx')
assert 'https://example.com' in part('out.pptx', 'ppt/slides/_rels/slide1.xml.rels')
