# area: headers-footers
# needs: #324
# rpptx's one-call footer, as an agent finds it in the rpptx docs
from rpptx import Presentation

prs = Presentation()
for title in ['One', 'Two']:
    prs.slides.add_slide(prs.slide_layouts[1]).shapes.title.text = title
prs.set_header_footer(slide_number=True, footer='Confidential')
prs.save('out.pptx')
# --- check
xml = part('out.pptx', 'ppt/slides/slide2.xml')
assert 'type="sldNum"' in xml and 'Confidential' in xml
