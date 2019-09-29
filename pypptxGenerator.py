


"""
code generator below should be in a different file
"""
def pptx2py():
	mergeMe_py = open("mergeMe.py","w+")
	writeImports(mergeMe_py)
	writeCreateNewPresentation(mergeMe_py)
	# detect && write if detected
	mergeMe_py.close()

def writeTable(__file__):
	__file__.write()

def eraseMergeMe_py():
	mergeMe_py = open("mergeMe.py","w+")
	mergeMe_py.close()

def writeImports(__file__):
	__file__.write("from pptx import Presentation\n\n")


def writeCreateNewPresentation(__file__):
	__file__.write("prs = Presentation()\n")
	__file__.write("title_slide_layout = prs.slide_layouts[0]\n")
	__file__.write("slide = prs.slides.add_slide(title_slide_layout)\n")
	__file__.write("title = slide.shapes.title\n")
	__file__.write("subtitle = slide.placeholders[1]\n")
	__file__.write("title.text = \"Hello, World!\"\n")
	__file__.write("subtitle.text = \"python-pptx generator was here!\"\n")
	__file__.write("prs.save('generated.pptx')\n")



if __name__ == "__main__":
	# generate file "mergeMe.py"
	pptx2py()


#find_table(prs)

"""



def write(__file__):
	__file__.write("\n")




"""
