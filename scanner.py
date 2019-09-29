import os
from pptx import Presentation
import generator
import re
import sys
import generator

"""
pptx content detector should be in its own file
"""





# def find_text(prs):
# 	for slide in prs.slides:
# 	    for shape in slide.shapes:
# 	        if not shape.has_text_frame:
# 	        	print("Slide " + str((prs.slides.index(slide)) + 1) + " has NO text.")
# 	        	continue
# 	        for paragraph in shape.text_frame.paragraphs:
# 	            for run in paragraph.runs:
# 	            	print("Slide " + str((prs.slides.index(slide)) + 1) + " has text:")
# 	            	# text_runs.append(run.text)
# 	            	print(run.text)


def find_table(prs, __file__):
	for slide in prs.slides:
		for shape in slide.shapes:
			try:
				if shape.table:
					print("Slide " + str((prs.slides.index(slide)) + 1) + " has table.")
					#writeTable(__file__)
					continue
			except:
				# print("Slide " + str((prs.slides.index(slide)) + 1) + " has NO table.")
				pass
			# for paragraph in shape.text_frame.paragraphs:
			#     for run in paragraph.runs:
			#     	print("Slide " + str((prs.slides.index(slide)) + 1) + " has text:")
			#     	# text_runs.append(run.text)
			#     	print(run.text)

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
						print("Slide " + str((prs.slides.index(slide)) + 1) + " has text paragraph: \n" + test_runs)
					#writeTable(__file__)
					continue
			except:
				# print("Slide " + str((prs.slides.index(slide)) + 1) + " has NO table.")
				pass

def string_found(string1, string2):
   if re.search(r"\b" + re.escape(string1) + r"\b", string2):
      return True
   return False

def findShapeDimensions(shape):
	print("\n TEXT_BOX LEFT: " + str(shape.left))
	print("\n TEXT_BOX TOP: " + str(shape.top))
	print("\n TEXT_BOX WIDTH: " + str(shape.width))
	print("\n TEXT_BOX HEIGHT: " + str(shape.height))


def find_placeholder(prs):
	for slide in prs.slides:
		for shape in slide.placeholders:
			try:
				if shape.placeholder_format:
					print("\n\n#############################################################\n\n")
					print("Slide " + str((prs.slides.index(slide)) + 1) + " has placeholder_format.")
					print("\nShape Type: " + str(shape.placeholder_format.type))

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
					if(string_found('TITLE',str(shape.placeholder_format.type))):
						print("\n FOUND A TITLE")
						if(shape.text_frame):
							print("\nTEXT_FRAME TEXT: " + shape.text_frame.text)
						else:
							print("\nTEXT_FRAME TEXT: NOT FOUND!!!")
					if(string_found('CENTER_TITLE',str(shape.placeholder_format.type))):
						print("\n FOUND A CENTER_TITLE")
						if(shape.text_frame):
							print("\nTEXT_FRAME TEXT: " + shape.text_frame.text)
						else:
							print("\nTEXT_FRAME TEXT: NOT FOUND!!!")
					if(string_found('SUBTITLE',str(shape.placeholder_format.type))):
						print("\n FOUND A SUBTITLE")
						if(shape.text_frame):
							print("\nTEXT_FRAME TEXT: " + shape.text_frame.text)
						else:
							print("\nTEXT_FRAME TEXT: NOT FOUND!!!")
					if(string_found('TABLE',str(shape.placeholder_format.type))):
						print("\n FOUND A TABLE")
					if(string_found('PICTURE',str(shape.placeholder_format.type))):
						print("\n FOUND A PICTURE")
					if(string_found('CHART',str(shape.placeholder_format.type))):
						print("\n FOUND A CHART")
					if(string_found('BODY',str(shape.placeholder_format.type))):
						print("\n FOUND A BODY")
						if(shape.text_frame):
							print("\nTEXT_FRAME TEXT: " + shape.text_frame.text)
						elif(shape.text):
							print("\nSHAPE.TEXT: " + shape.text)
						else:
							print("\nTEXT_FRAME TEXT: NOT FOUND!!!")
					if(string_found('OBJECT',str(shape.placeholder_format.type))):
						print("\n FOUND A OBJECT")
						if(shape.text_frame):
							print("\nTEXT_FRAME TEXT: " + shape.text_frame.text)
						else:
							print("\nTEXT_FRAME TEXT: NOT FOUND!!!")
					print('\tidx: %d \n\tname: %s\n\tplaceholder_format_type: %s\n\tshape.shape_type: %s' % (shape.placeholder_format.idx, shape.name, str(shape.placeholder_format.type), shape.shape_type))
					continue
			except Exception as detail:
				print(detail)
				pass


def find_shape(prs, GENERATED_FILE_NAME):
	for slide in prs.slides:
		generator.writeAddSlide(GENERATED_FILE_NAME)
		for shape in slide.shapes:
			try:
				print("\n\n#############################################################\n\n")
				print("Slide " + str((prs.slides.index(slide)) + 1) + " has a shape.")
				print("\nShape Type: " + str(shape.shape_type))

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
				if(string_found('PLACEHOLDER',str(shape.shape_type))):
					print("\n FOUND A PLACEHOLDER ON SHAPE")
					if(string_found('TITLE',str(shape.placeholder_format.type))):
						print("\n FOUND A TITLE")
						findShapeDimensions(shape)
						if(shape.text_frame):
							print("\nTEXT_FRAME TEXT: " + shape.text_frame.text)
						else:
							print("\nTEXT_FRAME TEXT: NOT FOUND!!!")

					if(string_found('CENTER_TITLE',str(shape.placeholder_format.type))):
						print("\n FOUND A CENTER_TITLE")
						findShapeDimensions(shape)

						if(shape.text_frame):
							print("\nTEXT_FRAME TEXT: " + shape.text_frame.text)
							generator.writeADD_CENTER_TITLE(GENERATED_FILE_NAME,shape.left,shape.top,shape.width,shape.height,shape.text_frame.text)
						else:
							print("\nTEXT_FRAME TEXT: NOT FOUND!!!")

					if(string_found('SUBTITLE',str(shape.placeholder_format.type))):
						print("\n FOUND A SUBTITLE")
						findShapeDimensions(shape)
						if(shape.text_frame):
							generator.writeADD_TEXTBOX(GENERATED_FILE_NAME,shape.left,shape.top,shape.width,shape.height,shape.text_frame.text)
							print("\nTEXT_FRAME TEXT: " + shape.text_frame.text)
						else:
							print("\nTEXT_FRAME TEXT: NOT FOUND!!!")

					if(string_found('TABLE',str(shape.placeholder_format.type))):
						print("\n FOUND A TABLE")

					if(string_found('PICTURE',str(shape.placeholder_format.type))):
						print("\n FOUND A PICTURE")

					if(string_found('CHART',str(shape.placeholder_format.type))):
						print("\n FOUND A CHART")

					if(string_found('BODY',str(shape.placeholder_format.type))):
						print("\n FOUND A BODY")
						findShapeDimensions(shape)
						if(shape.text_frame):
							print("\nTEXT_FRAME TEXT: " + shape.text_frame.text)
						elif(shape.text):
							print("\nSHAPE.TEXT: " + shape.text)
						else:
							print("\nTEXT_FRAME TEXT: NOT FOUND!!!")

					if(string_found('OBJECT',str(shape.placeholder_format.type))):
						print("\n FOUND A OBJECT")
						findShapeDimensions(shape)
						if(shape.text_frame):
							print("\nTEXT_FRAME TEXT: " + shape.text_frame.text)
						else:
							print("\nTEXT_FRAME TEXT: NOT FOUND!!!")
					print('\tidx: %d \n\tname: %s\n\tplaceholder_format_type: %s\n\tshape.shape_type: %s' % (shape.placeholder_format.idx, shape.name, str(shape.placeholder_format.type), shape.shape_type))
				else:
					print("\n NEW SHAPE_TYPE: " + str(shape.shape_type))
					if(string_found('TEXT_BOX',str(shape.shape_type))):
						print("\n FOUND A TEXT_BOX")
						findShapeDimensions(shape)
						if(shape.text):
							generator.writeADD_TEXTBOX(GENERATED_FILE_NAME,shape.left,shape.top,shape.width,shape.height,shapt.text_frame.text)
							print("\nSHAPE TEXT: " + shape.text)
						else:
							print("\nSHAPE TEXT: NOT FOUND!!!")
				continue
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


def scanPresentationByMethod(FILE_NAME): # FILE_NAME EXPECTS FILENAME (NO QUOTES, OR PREFIX)
	try: # try if ran at command line with file as first arg
		current_dr = os.path.dirname(os.path.realpath(__file__))
		FULL_PATH = current_dr + "/" + FILE_NAME 
		print(FULL_PATH)
		prs = Presentation(FULL_PATH)
		return prs
	except Exception as detail:
		print(detail)
		pass		
		

if __name__ == "__main__":
	prs = scanPresentationByCLI()

	## generate file "mergeMe.py"
	# pptx2py()	    		
	# current_dr = os.path.dirname(os.path.realpath(__file__))
	# prs = Presentation(current_dr + '/generated_user1.pptx')
	mergeMe_py = open("mergeMe.py","w+")
	find_shape(prs, mergeMe_py)
	#mergeMe_py = open("mergeMe.py","w+")
	#generator.writeImports(mergeMe_py)

def scanAndGenerate(prs, GENERATED_FILE_NAME):
	mergeMe_py = generator.openMergeMe_py(GENERATED_FILE_NAME)	
	generator.writeImports(mergeMe_py)
	generator.writeCreateNewPresentation(mergeMe_py)
	find_shape(prs, mergeMe_py)
	generator.writeSave(mergeMe_py)
	mergeMe_py.close()
