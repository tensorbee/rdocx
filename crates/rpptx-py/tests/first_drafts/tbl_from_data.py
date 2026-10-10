# area: tables
# needs: #326
from rpptx import Presentation
from rpptx.util import Inches

data = [['Name', 'Score'], ['Ann', '91'], ['Bob', '78']]
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
slide.shapes.title.text = 'Scores'
graphic_frame = slide.shapes.add_table(len(data), 2, Inches(1), Inches(2), Inches(4), Inches(1.2))
table = graphic_frame.table
for r, row in enumerate(data):
    for c, value in enumerate(row):
        table.cell(r, c).text = value
prs.save('out.pptx')
# --- check
xml = part('out.pptx')
assert all(v in xml for row in data for v in row)
