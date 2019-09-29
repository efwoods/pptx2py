


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

def writeTable(__file__):
	__file__.write()

def eraseMergeMe_py():
	mergeMe_py = open("mergeMe.py","w+")
	mergeMe_py.close()

def openMergeMe_py(__file__):
	mergeMe_py = open(__file__,"w+")
	return mergeMe_py

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

def writeADD_TEXTBOX(__file__, LEFT, TOP, WIDTH, HEIGHT, TEXT):
	__file__.write("shape = slide.shapes.add_textbox(" 
	+ str(LEFT) + "," + str(TOP) + "," + str(WIDTH) + "," + str(HEIGHT) + ")\n") #left top width height
	__file__.write("shape.text_frame.text = \"" + TEXT + "\"\n")

def writeADD_CENTER_TITLE(__file__, LEFT, TOP, WIDTH, HEIGHT, TEXT):
	__file__.write("shape = slide.shapes.add_textbox(" 
	+ str(LEFT) + "," + str(TOP) + "," + str(WIDTH) + "," + str(HEIGHT) + ")\n") #left top width height
	__file__.write("shape.text_frame.text = \"" + TEXT + "\"\n")



def writeTEXT(__file__): # BE MINDFUL OF SHAPE BEFORE USE !!!!!!!!!!!!
	__file__.write("shape.text_frame.text = \"ADDED TEXT HERE! :)\"\n")	


def writeSave(__file__):
	__file__.write("prs.save('generated_user1.pptx')\n")



if __name__ == "__main__":
	# generate file "mergeMe.py"
#	pptx2py()

"""



def write(__file__):
	__file__.write("\n")




"""
