import scanner
import sys
import generate as generator


def scanAndGenerate(prs, GENERATED_FILE_NAME):
	mergeMe_py = generator.openMergeMe_py(GENERATED_FILE_NAME)	
	generator.writeImports(mergeMe_py)
	generator.writeCreateNewPresentation(mergeMe_py)
	loopThroughPresentation(prs, mergeMe_py)
	
	generator.writeSave(mergeMe_py)
	generator.closeMergeMe_py(mergeMe_py)

def test(prs, GENERATED_FILE_NAME, OUTPUT_PPTX_FILENAME):
	mergeMe_py = generator.openMergeMe_py(GENERATED_FILE_NAME)
	generator.writeImports(mergeMe_py)
	generator.writeCreateNewPresentation(mergeMe_py)
	loopThroughPresentation(prs, mergeMe_py)
	
	generator.writeSave(mergeMe_py,OUTPUT_PPTX_FILENAME)
	generator.closeMergeMe_py(mergeMe_py)

def loopThroughPresentation(prs, GENERATED_FILE_HANDLER):
	for slide in prs.slides:
		generator.writeAddSlide(GENERATED_FILE_HANDLER)
		print("\n\n#############################################################\n\n")
		print("Slide " + str((prs.slides.index(slide)) + 1) + " has the following shapes")
		for shape in slide.shapes:
			try:
				identifyShapeAndGenerate(shape, GENERATED_FILE_HANDLER)
				continue
			except Exception as detail:
				print(detail)
				pass

def identifyShapeAndGenerate(shape, GENERATED_FILE_HANDLER):
			try:
				print("\nShape Type: " + str(shape.shape_type))
				""" print("\nShape.title.text: " + str(shape.title.text)) """
				if not (findPlaceholderAndGenerate(shape, GENERATED_FILE_HANDLER)):
					findShapeAndGenerate(shape, GENERATED_FILE_HANDLER)
			except Exception as detail:
				print(detail)
				pass

def findPlaceholderAndGenerate(shape, GENERATED_FILE_HANDLER):
	if(scanner.string_found('PLACEHOLDER',str(shape.shape_type))):
		print("\n FOUND A PLACEHOLDER ON SHAPE")
		FOUND_BY = 'BY PLACEHOLDER'
		if(scanner.findPlaceholderTITLE(shape)):
			scanner.separateFoundShapesWithPrint(FOUND_BY)

		elif(scanner.findPlaceholderCENTER_TITLE(shape)):
			scanner.separateFoundShapesWithPrint(FOUND_BY)

		elif(scanner.findPlaceholderSUBTITLE(shape)):
			scanner.separateFoundShapesWithPrint(FOUND_BY)

		elif(scanner.findPlaceholderTABLE(shape)):
			scanner.separateFoundShapesWithPrint(FOUND_BY)

		elif(scanner.findPlaceholderPICTURE(shape)):
			scanner.separateFoundShapesWithPrint(FOUND_BY)

		elif(scanner.findPlaceholderCHART(shape)):
			scanner.separateFoundShapesWithPrint(FOUND_BY)

		elif(scanner.findPlaceholderBODY(shape)):
			scanner.separateFoundShapesWithPrint(FOUND_BY)
			
		elif(scanner.findPlaceholderOBJECT(shape)):
			scanner.separateFoundShapesWithPrint(FOUND_BY)
		else:
			print('\n NO SHAPE IDENTIFIED BY PLACEHOLDER\n')
			return False
		return True
	return False

def findShapeAndGenerate(shape, GENERATED_FILE_HANDLER):
	print("\n SEARCHING FOR SHAPE BY SHAPE.SHAPE_TYPE: " + str(shape.shape_type))
	FOUND_BY = 'BY SHAPE.SHAPE_TYPE'
	
	if(scanner.findShapeTEXT_BOX(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
		generator.writeADD_TEXTBOX(GENERATED_FILE_HANDLER, shape)

	elif(scanner.findShapeAUTO_SHAPE(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeCALLOUT(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeCANVAS(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeCHART(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeCOMMENT(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeDIAGRAM(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeEMBEDDED_OLE_OBJECT(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeFORM_CONTROL(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeFREEFORM(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeGROUP(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeIGX_GRAPHIC(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeINK(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeINK_COMMENT(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeLINE(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeLINKED_OLE_OBJECT(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeLINKED_PICTURE(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeMEDIA(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeOLE_CONTROL_OBJECT(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapePICTURE(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapePLACEHOLDER(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeSCRIPT_ANCHOR(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)

	elif(scanner.findShapeTABLE(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
		generator.writeTable(GENERATED_FILE_HANDLER, shape)

	elif(scanner.findShapeTEXT_EFFECT(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeWEB_VIDEO(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	elif(scanner.findShapeMIXED(shape)):
		scanner.separateFoundShapesWithPrint(FOUND_BY)
	else:
		print("\n NEW SHAPE_TYPE: " + str(shape.shape_type))
		scanner.separateFoundShapesWithPrint(FOUND_BY)


if __name__ == "__main__":
	FILE_NAME = sys.argv[1]#"template_sample.pptx"
	GENERATED_FILE_NAME = "mergeMe.py"
	OUTPUT_PPTX_FILENAME = 'FINAL.pptx'

	prs = scanner.scanPresentationByMethod(FILE_NAME)

	#scanAndGenerate(prs,GENERATED_FILE_NAME)
	test(prs, GENERATED_FILE_NAME, OUTPUT_PPTX_FILENAME)