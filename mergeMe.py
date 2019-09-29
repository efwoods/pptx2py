from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
shape = slide.shapes.add_textbox(640080,5029200,1920240,346320)
shape.text_frame.text = "User2 CHANGE"
prs.save('generated_user1.pptx')
