import scanner

if __name__ == "__main__":
	## generate file "mergeMe.py"
	# pptx2py()	    
	FILE_NAME = "template_sample"		
	prs = scanner.scanPresentationByMethod(FILE_NAME)
	scanner.find_shape(prs)

	#mergeMe_py = open("mergeMe.py","w+")
	#generator.writeImports(mergeMe_py)
	