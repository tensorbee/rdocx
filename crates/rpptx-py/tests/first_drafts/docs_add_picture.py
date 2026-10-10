# area: docs
# needs: #326
# python-pptx getting started, add_picture() example
from rpptx import Presentation
from rpptx.util import Inches

img_path = 'monty-truth.png'

prs = Presentation()
blank_slide_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_slide_layout)

left = top = Inches(1)
pic = slide.shapes.add_picture(img_path, left, top)

left = Inches(5)
height = Inches(5.5)
pic = slide.shapes.add_picture(img_path, left, top, height=height)

prs.save('test.pptx')
# --- check
xml = part('test.pptx')
assert xml.count('<p:pic>') == 2 and 'cy="5029200"' in xml
