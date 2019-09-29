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
	