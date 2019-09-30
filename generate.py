from io import StringIO
import sys

"""
############################################################################
UTILITY FUNCTIONS
############################################################################
"""

def string_found(string1, string2):
   if re.search(r"\b" + re.escape(string1) + r"\b", string2):
      return True
   return False


def hex2RGB(hex, __file__):
	print('XXXXX XXXXXXXX 	 '+ hex+'		 XXXXXXXXXXXXXX')
	old_stdout = sys.stdout
	sys.stdout = __file__
	print('RGBColor', tuple(int(hex[i:i+2],16)for i in (0,2,4)), sep = '')
	sys.stdout = old_stdout


"""
code generator below should be in a different file
"""
def pptx2py():
	mergeMe_py = open("mergeMe.py","w+")
	writeImports(mergeMe_py)
	writeCreateNewPresentation(mergeMe_py)
	writeAddSlide(mergeMe_py)
	writeTITLE(mergeMe_py)
	writeSUBTITLE(mergeMe_py)
	writeADD_AUTO_SHAPE(mergeMe_py)
	writeTEXT(mergeMe_py)
	writeADD_TEXTBOX(mergeMe_py)
	writeSave(mergeMe_py)


	# detect && write if detected
	mergeMe_py.close()

def writeTable(__file__, shape):
	
	try:
		LEFT = str(shape.left)
		TOP = str(shape.top)
		WIDTH = str(shape.width)
		HEIGHT = str(shape.height)
		ROWS = str(len(shape.table.rows))
		COLS = str(len(shape.table.columns))

		__file__.write("LEFT = " + LEFT + "\n")
		__file__.write("TOP = " + TOP + "\n")
		__file__.write("WIDTH = " + WIDTH + "\n")
		__file__.write("HEIGHT = " + HEIGHT + "\n")
		__file__.write("ROWS = " + ROWS + "\n")
		__file__.write("COLS = " + COLS + "\n")

		__file__.write("shape = slide.shapes.add_table(ROWS, COLS, LEFT, TOP, WIDTH, HEIGHT)") #.TABLE

		__file__.write("\nfor row in range (0, len(shape.table.rows)):")
		__file__.write("\n\tfor col in range (0,len(shape.table.columns)):")
		__file__.write("\n\t\tif(row == 0):")
		for row in range (0, len(shape.table.rows)):
			for col in range (0, len(shape.table.columns)):
				if(row == 0):
					__file__.write("\n\t\t\tshape.table.columns[" + str(col) +"].width = " + str(shape.table.columns[col].width))

					__file__.write("\n\t\t\tshape.table.cell(" + str(row) + "," + str(col) + ").text = '" + str(shape.table.cell(row,col).text) + "'")
					__file__.write("\n\t\t\tshape.table.cell(" + str(row) + "," + str(col) + ").fill.solid()")
					""" __file__.write("\n\t\t\tshape.table.cell(" + str(row) + "," + str(col) + ").fill.type = '" + str(shape.table.cell(row,col).fill.type) + "'") """
					try:
						__file__.write("\n\t\t\tshape.table.cell(" + str(row) + "," + str(col) + ").fill.fore_color.rgb = ")
						hex2RGB(str(shape.table.cell(row,col).fill.fore_color.rgb),__file__)
						__file__.write("\n")
					except Exception as detail:
						print(detail)
						pass
				else:
					__file__.write("\n\t\tshape.table.cell(" + str(row) + "," + str(col) + ").text = '" + str(shape.table.cell(row,col).text) + "'")
					__file__.write("\n\t\tshape.table.cell(" + str(row) + "," + str(col) + ").fill.solid()")
					""" __file__.write("\n\t\tshape.table.cell(" + str(row) + "," + str(col) + ").fill.type = '" + str(shape.table.cell(row,col).fill.type) + "'") """
					
					try:
						__file__.write("\n\t\tshape.table.cell(" + str(row) + "," + str(col) + ").fill.fore_color.rgb = ")
						hex2RGB(str(shape.table.cell(row,col).fill.fore_color.rgb),__file__)
						__file__.write("\n")
					except Exception as detail:
						print(detail)
						pass
				

		__file__.write("\n")
	except Exception as detail:
		print(detail)
		pass
			
def eraseMergeMe_py():
	mergeMe_py = open("mergeMe.py","w+")
	mergeMe_py.close()

def openMergeMe_py(__file__):
	mergeMe_py = open(__file__,"w+")
	return mergeMe_py

def closeMergeMe_py(__file__):
	__file__.close()

def writeImports(__file__):
	__file__.write("from pptx import Presentation\n")
	__file__.write("from pptx.enum.shapes import MSO_SHAPE\n")
	__file__.write("from pptx.util import Inches, Pt\n")
	__file__.write("from pptx.dml.color import RGBColor\n")
	__file__.write("\n")

def writeCreateNewPresentation(__file__):
	__file__.write("prs = Presentation()\n")
	__file__.write("title_slide_layout = prs.slide_layouts[0]\n")

def writeAddSlide(__file__):
	__file__.write("slide = prs.slides.add_slide(title_slide_layout)\n")

def writeTITLE(__file__):
	__file__.write("title = slide.shapes.title\n")
	__file__.write("title.text = \"Hello, World!\"\n")

def writeSUBTITLE(__file__):
	__file__.write("subtitle = slide.placeholders[1]\n") ## placeholders[1]??
	__file__.write("subtitle.text = \"python-pptx generator was here!\"\n")

def writeADD_AUTO_SHAPE(__file__):
	__file__.write("left = Inches(1.0)\n")
	__file__.write("top = Inches(1.0)\n")
	__file__.write("width = Inches(1.0)\n")
	__file__.write("height = Inches(1.0)\n")
	__file__.write("shape = slide.shapes.add_shape(\n\tMSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height\n)\n")

def writeFILL_SHAPE_SOLID_RED(__file__):
	writeADD_AUTO_SHAPE(__file__)
	__file__.write("fill = shape.fill\n")
	__file__.write("fill.solid()\n")
	__file__.write("fill.fore_color.rgb = RGBColor(255, 0, 0)\n")

def writeADD_TEXTBOX(__file__, shape):
	LEFT = str(shape.left)
	TOP = str(shape.top)
	WIDTH = str(shape.width)
	HEIGHT = str(shape.height)

	__file__.write("LEFT = " + LEFT + "\n")
	__file__.write("TOP = " + TOP + "\n")
	__file__.write("WIDTH = " + WIDTH + "\n")
	__file__.write("HEIGHT = " + HEIGHT + "\n")

	__file__.write("shape = slide.shapes.add_textbox(LEFT,TOP,WIDTH,HEIGHT)\n") #left top width height
	

	if(shape.text_frame):
		text_runs = []
		""" __file__.write("\nfor paragraph in range (0, len(shape.text_frame.paragraphs)):")
		__file__.write("\n\tfor run in range (0,len(paragraph.runs)):")
		 """
		for paragraph in shape.text_frame.paragraphs:
			for run in paragraph.runs:
				text_runs.append(run.text)
		TEXT = ','.join(text_runs)

	else:
		TEXT = shape.text

	__file__.write("shape.text_frame.text = '" + TEXT + "'\n")


def writeADD_CENTER_TITLE(__file__, LEFT, TOP, WIDTH, HEIGHT, TEXT):
	__file__.write("shape = slide.shapes.add_textbox(" 
	+ str(LEFT) + "," + str(TOP) + "," + str(WIDTH) + "," + str(HEIGHT) + ")\n") #left top width height
	__file__.write("shape.text_frame.text = \"" + TEXT + "\"\n")



def writeTEXT(__file__): # BE MINDFUL OF SHAPE BEFORE USE !!!!!!!!!!!!
	__file__.write("shape.text_frame.text = \"ADDED TEXT HERE! :)\"\n")	


def writeSave(__file__, OUTPUT_PPTX_FILENAME):
	if(string_found('.pptx',OUTPUT_PPTX_FILENAME)):
		__file__.write("prs.save('" + OUTPUT_PPTX_FILENAME + "')\n")
	else:
		__file__.write("prs.save('" + OUTPUT_PPTX_FILENAME + ".pptx')\n")
	


"""

if __name__ == "__main__":
	# generate file "mergeMe.py"
#	pptx2py()




def write(__file__):
	__file__.write("\n")




"""
