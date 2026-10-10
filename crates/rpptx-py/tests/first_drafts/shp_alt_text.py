# area: shapes
# needs: #326
from rpptx import Presentation
from rpptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
pic = slide.shapes.add_picture('logo.png', Inches(1), Inches(1), width=Inches(1))
pic.alt_text = 'Company logo'
prs.save('out.pptx')
# --- check
assert 'descr="Company logo"' in part('out.pptx')
