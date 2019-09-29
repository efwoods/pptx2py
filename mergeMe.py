from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
LEFT = 685800
TOP = 2895480
WIDTH = 7772040
HEIGHT = 1066320
shape = slide.shapes.add_textbox(LEFT,TOP,WIDTH,HEIGHT)
shape.text_frame.text = 'Presentation Title'
LEFT = 685800
TOP = 4648320
WIDTH = 6400440
HEIGHT = 1904760
shape = slide.shapes.add_textbox(LEFT,TOP,WIDTH,HEIGHT)
shape.text_frame.text = 'Author,Department,Date,Location'
slide = prs.slides.add_slide(title_slide_layout)
LEFT = 722160
TOP = 2906640
WIDTH = 7772040
HEIGHT = 1361880
shape = slide.shapes.add_textbox(LEFT,TOP,WIDTH,HEIGHT)
shape.text_frame.text = 'Presentation Title'
LEFT = 722160
TOP = 4648320
WIDTH = 7772040
HEIGHT = 1683360
shape = slide.shapes.add_textbox(LEFT,TOP,WIDTH,HEIGHT)
shape.text_frame.text = 'Author,Department,Date,Location'
slide = prs.slides.add_slide(title_slide_layout)
slide = prs.slides.add_slide(title_slide_layout)
LEFT = 533520
TOP = 990720
WIDTH = 3007800
HEIGHT = 552240
shape = slide.shapes.add_textbox(LEFT,TOP,WIDTH,HEIGHT)
shape.text_frame.text = 'GraphTitle'
LEFT = 533520
TOP = 6019920
WIDTH = 8229240
HEIGHT = 685440
shape = slide.shapes.add_textbox(LEFT,TOP,WIDTH,HEIGHT)
shape.text_frame.text = 'Additional Notes ,- ,Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem'
slide = prs.slides.add_slide(title_slide_layout)
LEFT = 533520
TOP = 990720
WIDTH = 8152920
HEIGHT = 533160
shape = slide.shapes.add_textbox(LEFT,TOP,WIDTH,HEIGHT)
shape.text_frame.text = 'Chart Title'
LEFT = 533520
TOP = 1676520
WIDTH = 8152920
HEIGHT = 4419360
ROWS = 7
COLS = 4
shape = slide.shapes.add_table(ROWS, COLS, LEFT, TOP, WIDTH, HEIGHT)
for row in range (0, len(shape.table.rows)):
	for col in range (0,len(shape.table.columns)):
		if(row == 0):
			shape.table.columns[0].width = 2038320
			shape.table.cell(0,0).text = 'Column A'
			shape.table.cell(0,0).fill.solid()
			shape.table.cell(0,0).fill.fore_color.rgb = RGBColor(192, 0, 0)


			shape.table.columns[1].width = 2038320
			shape.table.cell(0,1).text = 'B'
			shape.table.cell(0,1).fill.solid()
			shape.table.cell(0,1).fill.fore_color.rgb = RGBColor(192, 0, 0)


			shape.table.columns[2].width = 2038320
			shape.table.cell(0,2).text = 'C'
			shape.table.cell(0,2).fill.solid()
			shape.table.cell(0,2).fill.fore_color.rgb = RGBColor(192, 0, 0)


			shape.table.columns[3].width = 2038320
			shape.table.cell(0,3).text = 'D'
			shape.table.cell(0,3).fill.solid()
			shape.table.cell(0,3).fill.fore_color.rgb = RGBColor(192, 0, 0)


		shape.table.cell(1,0).text = 'XXXXXXXX'
		shape.table.cell(1,0).fill.solid()
		shape.table.cell(1,0).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(1,1).text = 'XX'
		shape.table.cell(1,1).fill.solid()
		shape.table.cell(1,1).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(1,2).text = 'XX'
		shape.table.cell(1,2).fill.solid()
		shape.table.cell(1,2).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(1,3).text = 'XX'
		shape.table.cell(1,3).fill.solid()
		shape.table.cell(1,3).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(2,0).text = 'XXXXXXXX'
		shape.table.cell(2,0).fill.solid()
		shape.table.cell(2,0).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(2,1).text = 'XX'
		shape.table.cell(2,1).fill.solid()
		shape.table.cell(2,1).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(2,2).text = 'XX'
		shape.table.cell(2,2).fill.solid()
		shape.table.cell(2,2).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(2,3).text = 'XX'
		shape.table.cell(2,3).fill.solid()
		shape.table.cell(2,3).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(3,0).text = 'XXXXXXXX'
		shape.table.cell(3,0).fill.solid()
		shape.table.cell(3,0).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(3,1).text = 'XX'
		shape.table.cell(3,1).fill.solid()
		shape.table.cell(3,1).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(3,2).text = 'XX'
		shape.table.cell(3,2).fill.solid()
		shape.table.cell(3,2).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(3,3).text = 'XX'
		shape.table.cell(3,3).fill.solid()
		shape.table.cell(3,3).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(4,0).text = 'XXXXXXXX'
		shape.table.cell(4,0).fill.solid()
		shape.table.cell(4,0).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(4,1).text = 'XX'
		shape.table.cell(4,1).fill.solid()
		shape.table.cell(4,1).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(4,2).text = 'XX'
		shape.table.cell(4,2).fill.solid()
		shape.table.cell(4,2).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(4,3).text = 'XX'
		shape.table.cell(4,3).fill.solid()
		shape.table.cell(4,3).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(5,0).text = 'XXXXXXXX'
		shape.table.cell(5,0).fill.solid()
		shape.table.cell(5,0).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(5,1).text = 'XX'
		shape.table.cell(5,1).fill.solid()
		shape.table.cell(5,1).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(5,2).text = 'XX'
		shape.table.cell(5,2).fill.solid()
		shape.table.cell(5,2).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(5,3).text = 'XX'
		shape.table.cell(5,3).fill.solid()
		shape.table.cell(5,3).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(6,0).text = 'XXXXXXXX'
		shape.table.cell(6,0).fill.solid()
		shape.table.cell(6,0).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(6,1).text = 'XX'
		shape.table.cell(6,1).fill.solid()
		shape.table.cell(6,1).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(6,2).text = 'XX'
		shape.table.cell(6,2).fill.solid()
		shape.table.cell(6,2).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(6,3).text = 'XX'
		shape.table.cell(6,3).fill.solid()
		shape.table.cell(6,3).fill.fore_color.rgb = RGBColor(244, 231, 231)


slide = prs.slides.add_slide(title_slide_layout)
LEFT = 533520
TOP = 990720
WIDTH = 8152920
HEIGHT = 533160
shape = slide.shapes.add_textbox(LEFT,TOP,WIDTH,HEIGHT)
shape.text_frame.text = 'List Title'
LEFT = 533520
TOP = 1676520
WIDTH = 8152920
HEIGHT = 4449240
shape = slide.shapes.add_textbox(LEFT,TOP,WIDTH,HEIGHT)
shape.text_frame.text = 'Lorem ipsum dolor sit amet ,Aenean commodo ligula eget dolor ,Cum sociis natoque penatibus et magnis dis parturient montes ,Donec quam felis, ultricies nec, pellentesque eu ,Lorem ipsum dolor sit amet, consectetuer adipiscing elit ,Aenean massa ,Aenean commodo ligula eget dolor'
slide = prs.slides.add_slide(title_slide_layout)
LEFT = 533520
TOP = 990720
WIDTH = 8152920
HEIGHT = 533160
shape = slide.shapes.add_textbox(LEFT,TOP,WIDTH,HEIGHT)
shape.text_frame.text = 'Slide Title'
LEFT = 533520
TOP = 1676520
WIDTH = 8152920
HEIGHT = 4449240
shape = slide.shapes.add_textbox(LEFT,TOP,WIDTH,HEIGHT)
shape.text_frame.text = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem.,Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem.'
slide = prs.slides.add_slide(title_slide_layout)
LEFT = 533520
TOP = 6248520
WIDTH = 8152920
HEIGHT = 347400
shape = slide.shapes.add_textbox(LEFT,TOP,WIDTH,HEIGHT)
shape.text_frame.text = 'Type image caption here.'
LEFT = 533520
TOP = 1066680
WIDTH = 8152920
HEIGHT = 456840
shape = slide.shapes.add_textbox(LEFT,TOP,WIDTH,HEIGHT)
shape.text_frame.text = 'Slide Title'
slide = prs.slides.add_slide(title_slide_layout)
LEFT = 507960
TOP = 1038240
WIDTH = 8229240
HEIGHT = 532440
shape = slide.shapes.add_textbox(LEFT,TOP,WIDTH,HEIGHT)
shape.text_frame.text = 'Page Title'
LEFT = 2834640
TOP = 2743200
WIDTH = 3749040
HEIGHT = 346320
shape = slide.shapes.add_textbox(LEFT,TOP,WIDTH,HEIGHT)
shape.text_frame.text = 'USER UPDATES WITH TEXTBOX'
prs.save('generated_FINAL.pptx')
