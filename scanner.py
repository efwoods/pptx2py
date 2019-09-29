"""
############################################################################
IMPORTS
############################################################################
"""

import os
from pptx import Presentation
import re
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

def findShapeDimensions(shape):
	print("\n TEXT_BOX LEFT: " + str(shape.left))
	print("\n TEXT_BOX TOP: " + str(shape.top))
	print("\n TEXT_BOX WIDTH: " + str(shape.width))
	print("\n TEXT_BOX HEIGHT: " + str(shape.height))

"""
############################################################################
FIND BY SHAPE_TYPE FUNCTIONS
############################################################################
"""

def findShapeTEXT_BOX(shape):
	if(string_found('TEXT_BOX',str(shape.shape_type))):
		print("\n FOUND A TEXT_BOX")
		findShapeDimensions(shape)
		if(shape.text):
			print("\nSHAPE TEXT: " + shape.text)
			return True
		else:
			print("\nSHAPE TEXT: NOT FOUND!!!")
	return False


"""
############################################################################
FIND BY PLACEHOLDER FUNCTIONS
############################################################################
"""

def findPlaceholder(shape):
	if(string_found('PLACEHOLDER',str(shape.shape_type))):
		print("\n FOUND A PLACEHOLDER ON SHAPE")
		return True
	return False

def findPlaceholderTITLE(shape):
	if(string_found('TITLE',str(shape.placeholder_format.type))):
		print("\n FOUND A TITLE")
		findShapeDimensions(shape)
		if(shape.text_frame):
			print("\nTEXT_FRAME TEXT: " + shape.text_frame.text)
			return True
		else:
			print("\nTEXT_FRAME TEXT: NOT FOUND!!!")
	return False

def findPlaceholderCENTER_TITLE(shape):
	if(string_found('CENTER_TITLE',str(shape.placeholder_format.type))):
		print("\n FOUND A CENTER_TITLE")
		findShapeDimensions(shape)
		if(shape.text_frame):
			print("\nTEXT_FRAME TEXT: " + shape.text_frame.text)
			return True
		else:
			print("\nTEXT_FRAME TEXT: NOT FOUND!!!")
	return False

def findPlaceholderSUBTITLE(shape):
	if(string_found('SUBTITLE',str(shape.placeholder_format.type))):
		print("\n FOUND A SUBTITLE")
		findShapeDimensions(shape)
		if(shape.text_frame):
			print("\nTEXT_FRAME TEXT: " + shape.text_frame.text)
			return True
		else:
			print("\nTEXT_FRAME TEXT: NOT FOUND!!!")
	return False

def findPlaceholderTABLE(shape):
	if(string_found('TABLE',str(shape.placeholder_format.type))):
		print("\n FOUND A TABLE")
		return True
	return False

def findPlaceholderPICTURE(shape):
	if(string_found('PICTURE',str(shape.placeholder_format.type))):
		print("\n FOUND A PICTURE")
		return True
	return False

def findPlaceholderCHART(shape):
	if(string_found('CHART',str(shape.placeholder_format.type))):
		print("\n FOUND A CHART")
		return True
	return False

def findPlaceholderBODY(shape):
	if(string_found('BODY',str(shape.placeholder_format.type))):
		print("\n FOUND A BODY")
		findShapeDimensions(shape)
		if(shape.text_frame):
			print("\nTEXT_FRAME TEXT: " + shape.text_frame.text)
			return True
		elif(shape.text):
			print("\nSHAPE.TEXT: " + shape.text)
			return True
		else:
			print("\nTEXT_FRAME TEXT: NOT FOUND!!!")
	return False

def findPlaceholderOBJECT(shape):
	if(string_found('OBJECT',str(shape.placeholder_format.type))):
		print("\n FOUND A OBJECT")
		findShapeDimensions(shape)
		if(shape.text_frame):
			print("\nTEXT_FRAME TEXT: " + shape.text_frame.text)
			return True
		else:
			print("\nTEXT_FRAME TEXT: NOT FOUND!!!")
	return False

"""
############################################################################
CORE FUNCTIONS
############################################################################
"""

def identifyAllShapes(prs):
	for slide in prs.slides:
		print("\n\n#############################################################\n\n")
		print("Slide " + str((prs.slides.index(slide)) + 1) + " has the following shapes:")
		for shape in slide.shapes:
			identifyShape(shape)

def identifyShape(shape):
			try:
				print("\nShape Type: " + str(shape.shape_type))
				try:
					print('\tidx: %d \n\tname: %s\n\tplaceholder_format_type: %s\n\tshape.shape_type: %s' % (shape.placeholder_format.idx, shape.name, str(shape.placeholder_format.type), shape.shape_type))
				except Exception as detail:
					print("Error on PLACEHOLDER LOG PRINT")
					print(detail)
					pass
				if(findPlaceholder(shape)):
					if(findPlaceholderTITLE(shape)):
						print('\n...EXITING ID SHAPE BY PLACEHOLDER...\n')

					elif(findPlaceholderCENTER_TITLE(shape)):
						print('\n...EXITING ID SHAPE...\n')

					elif(findPlaceholderSUBTITLE(shape)):
						print('\n...EXITING ID SHAPE BY PLACEHOLDER...\n')

					elif(findPlaceholderTABLE(shape)):
						print('\n...EXITING ID SHAPE BY PLACEHOLDER...\n')

					elif(findPlaceholderPICTURE(shape)):
						print('\n...EXITING ID SHAPE BY PLACEHOLDER...\n')

					elif(findPlaceholderCHART(shape)):
						print('\n...EXITING ID SHAPE BY PLACEHOLDER...\n')

					elif(findPlaceholderBODY(shape)):
						print('\n...EXITING ID SHAPE BY PLACEHOLDER...\n')
						
					elif(findPlaceholderOBJECT(shape)):
						print('\n...EXITING ID SHAPE BY PLACEHOLDER...\n')
					else:
						print('\n NO SHAPE IDENTIFIED BY PLACEHOLDER\n')
				else:
					print("\n SEARCHING FOR SHAPE BY SHAPE.SHAPE_TYPE: " + str(shape.shape_type))
					if(findShapeTEXT_BOX(shape)):
						print('\n...EXITING ID SHAPE BY SHAPE.SHAPE_TYPE...\n')
					else:
						print("\n NEW SHAPE_TYPE: " + str(shape.shape_type))
						print('\n...EXITING ID SHAPE BY SHAPE.SHAPE_TYPE...\n')
			except Exception as detail:
				print(detail)
				pass

def scanPresentationByCLI():
	try: # try if ran at command line with file as first arg
		FILE_NAME = sys.argv[1] # ABSOLUTE FILE PATH OF PPTX TO BE SCANNED
		current_dr = os.path.dirname(os.path.realpath(__file__))
		FULL_PATH = current_dr + "/" + FILE_NAME 
		print(FULL_PATH)
		prs = Presentation(FULL_PATH)
		return prs
	except Exception as detail:
		print(detail)
		pass	

def scanPresentationByMethod(FILE_NAME): # FILE_NAME EXPECTS FILENAME for Scanning (NO QUOTES, OR PREFIX)
	try: # try if ran at command line with file as first arg
		current_dr = os.path.dirname(os.path.realpath(__file__))
		FULL_PATH = current_dr + "/" + FILE_NAME 
		print(FULL_PATH)
		prs = Presentation(FULL_PATH)
		return prs
	except Exception as detail:
		print(detail)
		pass		

"""
############################################################################
MAIN
############################################################################
"""

if __name__ == "__main__":
	prs = scanPresentationByCLI()
	identifyAllShapes(prs)
"""
############################################################################
TEMPLATES
############################################################################
"""

""" def TEMPLATE TYPES TO CHECK
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

""" def TEMPLATE_FUNCTION
def find_text(prs):
	for slide in prs.slides:
		for shape in slide.shapes:
			try:
				if shape.text_frame:
					print("Slide " + str((prs.slides.index(slide)) + 1) + " has text_frame.")
					for paragraph in shape.text_frame.paragraphs:
						text_runs = []
						for run in paragraph.runs:
							text_runs.append(run.text)
							print(run.text)
						print("Slide " + str((prs.slides.index(slide)) + 1) + " has text paragraph: \n" + text_runs)
					#writeTable(__file__)
					continue
			except:
				# print("Slide " + str((prs.slides.index(slide)) + 1) + " has NO table.")
				pass
"""

""" def TEMPLATE_FUNCTION
def findShapeXXX(shape):
	if(
string_found('PLACEHOLDER',str(shape.shape_type))
):
		# CONTENT
		return True
	return False
"""
