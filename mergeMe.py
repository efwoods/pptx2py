from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
slide = prs.slides.add_slide(title_slide_layout)
slide = prs.slides.add_slide(title_slide_layout)
slide = prs.slides.add_slide(title_slide_layout)
slide = prs.slides.add_slide(title_slide_layout)
LEFT = 533520
TOP = 1676520
WIDTH = 8152920
HEIGHT = 4419000
ROWS = 8
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


		shape.table.cell(1,0).text = ' Alex, Evan'
		shape.table.cell(1,0).fill.solid()
		shape.table.cell(1,0).fill.fore_color.rgb = RGBColor(0, 166, 93)


		shape.table.cell(1,1).text = 'Due Oct 26th'
		shape.table.cell(1,1).fill.solid()
		shape.table.cell(1,1).fill.fore_color.rgb = RGBColor(192, 0, 0)


		shape.table.cell(1,2).text = ' These are '
		shape.table.cell(1,2).fill.solid()
		shape.table.cell(1,2).fill.fore_color.rgb = RGBColor(192, 0, 0)


		shape.table.cell(1,3).text = 'updates'
		shape.table.cell(1,3).fill.solid()
		shape.table.cell(1,3).fill.fore_color.rgb = RGBColor(192, 0, 0)


		shape.table.cell(2,0).text = 'XXXXXXXX'
		shape.table.cell(2,0).fill.solid()
		shape.table.cell(2,0).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(2,1).text = 'XX'
		shape.table.cell(2,1).fill.solid()
		shape.table.cell(2,1).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(2,2).text = 'XX'
		shape.table.cell(2,2).fill.solid()
		shape.table.cell(2,2).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(2,3).text = 'XX'
		shape.table.cell(2,3).fill.solid()
		shape.table.cell(2,3).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(3,0).text = 'XXXXXXXX'
		shape.table.cell(3,0).fill.solid()
		shape.table.cell(3,0).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(3,1).text = 'XX'
		shape.table.cell(3,1).fill.solid()
		shape.table.cell(3,1).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(3,2).text = 'XX'
		shape.table.cell(3,2).fill.solid()
		shape.table.cell(3,2).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(3,3).text = 'XX'
		shape.table.cell(3,3).fill.solid()
		shape.table.cell(3,3).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(4,0).text = 'XXXXXXXX'
		shape.table.cell(4,0).fill.solid()
		shape.table.cell(4,0).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(4,1).text = 'XX'
		shape.table.cell(4,1).fill.solid()
		shape.table.cell(4,1).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(4,2).text = 'XX'
		shape.table.cell(4,2).fill.solid()
		shape.table.cell(4,2).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(4,3).text = 'XX'
		shape.table.cell(4,3).fill.solid()
		shape.table.cell(4,3).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(5,0).text = 'XXXXXXXX'
		shape.table.cell(5,0).fill.solid()
		shape.table.cell(5,0).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(5,1).text = 'XX'
		shape.table.cell(5,1).fill.solid()
		shape.table.cell(5,1).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(5,2).text = 'XX'
		shape.table.cell(5,2).fill.solid()
		shape.table.cell(5,2).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(5,3).text = 'XX'
		shape.table.cell(5,3).fill.solid()
		shape.table.cell(5,3).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(6,0).text = 'XXXXXXXX'
		shape.table.cell(6,0).fill.solid()
		shape.table.cell(6,0).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(6,1).text = 'XX'
		shape.table.cell(6,1).fill.solid()
		shape.table.cell(6,1).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(6,2).text = 'XX'
		shape.table.cell(6,2).fill.solid()
		shape.table.cell(6,2).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(6,3).text = 'XX'
		shape.table.cell(6,3).fill.solid()
		shape.table.cell(6,3).fill.fore_color.rgb = RGBColor(232, 204, 204)


		shape.table.cell(7,0).text = 'XXXXXXXX'
		shape.table.cell(7,0).fill.solid()
		shape.table.cell(7,0).fill.fore_color.rgb = RGBColor(0, 102, 179)


		shape.table.cell(7,1).text = 'XX'
		shape.table.cell(7,1).fill.solid()
		shape.table.cell(7,1).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(7,2).text = 'XX'
		shape.table.cell(7,2).fill.solid()
		shape.table.cell(7,2).fill.fore_color.rgb = RGBColor(244, 231, 231)


		shape.table.cell(7,3).text = 'XX'
		shape.table.cell(7,3).fill.solid()
		shape.table.cell(7,3).fill.fore_color.rgb = RGBColor(244, 231, 231)


slide = prs.slides.add_slide(title_slide_layout)
slide = prs.slides.add_slide(title_slide_layout)
slide = prs.slides.add_slide(title_slide_layout)
slide = prs.slides.add_slide(title_slide_layout)
prs.save('generated_FINAL.pptx')
