from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
shape = slide.shapes.add_textbox(685800,2130425,7772400,1470025)
shape.text_frame.text = "Hello, World!"
shape = slide.shapes.add_textbox(1371600,3886200,6400800,1752600)
shape.text_frame.text = "python-pptx generator was here!"
prs.save('generated_user1.pptx')
