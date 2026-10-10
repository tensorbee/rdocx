# area: templates
# needs: #318
from rpptx import Presentation

prs = Presentation('template.pptx')
prs.core_properties.title = 'Quarterly review'
prs.core_properties.author = 'Finance team'
prs.save('out.pptx')
# --- check
core = part('out.pptx', 'docProps/core.xml')
assert 'Quarterly review' in core and 'Finance team' in core
