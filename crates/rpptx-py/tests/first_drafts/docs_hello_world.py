# area: docs
# needs: #326
# python-pptx getting started, "Hello World!" example
from rpptx import Presentation

prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]

title.text = "Hello, World!"
subtitle.text = "python-pptx was here!"

prs.save('test.pptx')
# --- check
xml = part('test.pptx')
assert 'Hello, World!' in xml and 'python-pptx was here!' in xml
