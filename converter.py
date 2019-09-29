import scanner
import sys
import generate as generator

if __name__ == "__main__":
	## generate file "mergeMe.py"
	# pptx2py()	    


	FILE_NAME = "template_sample.pptx"
	GENERATED_FILE_NAME = "mergeMe.py"
			
	prs = scanner.scanPresentationByMethod(FILE_NAME)
	#mergeMe_py = generator.openMergeMe_py(GENERATED_FILE_NAME)	
	
	scanner.scanAndGenerate(prs,GENERATED_FILE_NAME)
	
	
	#generator.writeImports(mergeMe_py)
	
""" 	
	generator.writeCreateNewPresentation(mergeMe_py)
	generator.writeAddSlide(mergeMe_py)
	generator.writeTITLE(mergeMe_py)
	generator.writeADD_AUTO_SHAPE(mergeMe_py)
	generator.writeTEXT(mergeMe_py)
	generator.writeADD_TEXTBOX(mergeMe_py)
"""
	#scanner.find_shape(prs, mergeMe_py)
	
	#generator.writeSave(mergeMe_py)

	#mergeMe_py = open("mergeMe.py","w+")
	#generator.writeImports(mergeMe_py)
	

def scanAndGenerate(prs, GENERATED_FILE_NAME):
	mergeMe_py = generator.openMergeMe_py(GENERATED_FILE_NAME)	
	generator.writeImports(mergeMe_py)
	generator.writeCreateNewPresentation(mergeMe_py)
	loopThroughPresentation(prs, mergeMe_py)
	scanner.find_shape(prs, mergeMe_py)
	generator.writeSave(mergeMe_py)
	generator.closeMergeMe_py(mergeMe_py)


"""
				USE ALL CAPS
				BITMAP
				    Bitmap
				BODY
				    Body
				CENTER_TITLE
				    Center Title
				CHART
				    Chart
				DATE
				    Date
				FOOTER
				    Footer
				HEADER
				    Header
				MEDIA_CLIP
				    Media Clip
				OBJECT
				    Object
				ORG_CHART
				    Organization Chart
				PICTURE
				    Picture
				SLIDE_NUMBER
				    Slide Number
				SUBTITLE
				    Subtitle
				TABLE
				    Table
				TITLE
				    Title
				VERTICAL_BODY
				    Vertical Body
				VERTICAL_OBJECT
				    Vertical Object
				VERTICAL_TITLE
				    Vertical Title
				MIXED
"""

def loopThroughPresentation(prs, GENERATED_FILE_NAME):
	for slide in prs.slides:
		generator.writeAddSlide(GENERATED_FILE_NAME)
		print("\n\n#############################################################\n\n")
		print("Slide " + str((prs.slides.index(slide)) + 1) + " has the following shapes")
		for shape in slide.shapes:
			try:
				scanner.identifyShape(shape)
				continue
			except Exception as detail:
				print(detail)
				pass


