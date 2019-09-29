from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
title = slide.shapes.title
title.text = "Hello, World!"
subtitle = slide.placeholders[1]
subtitle.text = "python-pptx generator was here!"
left = Inches(1.0)
top = Inches(1.0)
width = Inches(1.0)
height = Inches(1.0)
shape = slide.shapes.add_shape(
	MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
)
shape.text_frame.text = "ADDED TEXT HERE! :)"
shape = slide.shapes.add_textbox(2,1,1,1)
shape.text_frame.text = "ADDED TEXT HERE! :)"
prs.save('generated_user1.pptx')
