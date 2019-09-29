from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
shape = slide.shapes.add_textbox(685800,2895480,7772040,1066320)
shape.text_frame.text = "Presentation Title"
shape = slide.shapes.add_textbox(685800,4648320,6400440,1904760)
shape.text_frame.text = "Author\nDepartment\nDate\nLocation\n"
shape = slide.shapes.add_textbox(4480560,2651760,3017520,346320)
shape.text_frame.text = "User1 test "
slide = prs.slides.add_slide(title_slide_layout)
shape = slide.shapes.add_textbox(722160,2906640,7772040,1361880)
shape.text_frame.text = "Presentation Title"
shape = slide.shapes.add_textbox(722160,4648320,7772040,1683360)
shape.text_frame.text = "AuthorAuthor \nDepartment\nDate\nLocation\n"
slide = prs.slides.add_slide(title_slide_layout)
slide = prs.slides.add_slide(title_slide_layout)
shape = slide.shapes.add_textbox(533520,990720,3007800,552240)
shape.text_frame.text = "GraphTitle"
shape = slide.shapes.add_textbox(533520,6019920,8229240,685440)
shape.text_frame.text = "Additional Notes - Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem"
slide = prs.slides.add_slide(title_slide_layout)
shape = slide.shapes.add_textbox(533520,990720,8152920,533160)
shape.text_frame.text = "Chart Title"
slide = prs.slides.add_slide(title_slide_layout)
shape = slide.shapes.add_textbox(533520,990720,8152920,533160)
shape.text_frame.text = "List Title"
shape = slide.shapes.add_textbox(533520,1676520,8152920,4449240)
shape.text_frame.text = "Lorem ipsum dolor sit ame"
slide = prs.slides.add_slide(title_slide_layout)
shape = slide.shapes.add_textbox(533520,990720,8152920,533160)
shape.text_frame.text = "Slide Title"
shape = slide.shapes.add_textbox(533520,1676520,8152920,4449240)
shape.text_frame.text = "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem."
slide = prs.slides.add_slide(title_slide_layout)
shape = slide.shapes.add_textbox(533520,6248520,8152920,347400)
shape.text_frame.text = "Type image caption here."
shape = slide.shapes.add_textbox(533520,1066680,8152920,456840)
shape.text_frame.text = "Slide Title"
slide = prs.slides.add_slide(title_slide_layout)
shape = slide.shapes.add_textbox(507960,1038240,8229240,532440)
shape.text_frame.text = "Page Title"
shape = slide.shapes.add_textbox(2834640,2743200,3749040,346320)
shape.text_frame.text = "USER UPDATES WITH TEXTBOX"
prs.save('generated_user1.pptx')
