from comtypes.client import CreateObject
import PyPDF2
import time
import xml.etree.ElementTree as ET
import re
import argparse

# https://www.w3.org/TR/speech-synthesis/
# TODO: Add way to easily modify the voice 
class SSML_Generator:
	def __init__(self,pause):
		self.pause = pause
		
	def __call__(self,text):
		SSML_document = self._header()
		SSML_document.text = text[0]
		for utterance in text[1:]:
			brk = self._pause(SSML_document)
			brk.tail = utterance
		return SSML_document
		
	def _pause(self,parent_tag):
		return ET.SubElement(parent_tag,"break",{"time":str(self.pause)+"ms"})

	def _header(self):
		return ET.Element("speak",{"version":"1.0", "xmlns":"http://www.w3.org/2001/10/synthesis", "xml:lang":"en-US"})

	# Nice to have: Transform acronyms into their pronunciation (See say as tag)
	
class SpeechToWav:
	def __init__(self,pause,engineRate):
		self.engine = CreateObject("SAPI.SpVoice")
		self.stream = CreateObject("SAPI.SpFileStream")
		self.SSML_generator = SSML_Generator(pause)
		# This import has to be done in a specific order, otherwise Windows is not happy
		from comtypes.gen import SpeechLib
		# Because of that, below is a work around to use its functions throughout the object.
		self.SpeechLib = SpeechLib
		self.engine.Rate = engineRate
		
	def __call__(self,text,outputFile):
		# https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms723606(v%3Dvs.85)
		self.stream.Open(outputFile + ".wav", self.SpeechLib.SSFMCreateForWrite)
		self.engine.AudioOutputStream = self.stream
		text = self._text_processing(text)
		text = self.SSML_generator(text)
		text = ET.tostring(text,encoding='utf8', method='xml').decode('utf-8')
		# Replace with SSML file and flag
		self.engine.speak(text)
		self.stream.Close()
	
	
	def _text_processing(self,text):
		""" Used to remove side effects from columns, the new lines add bizzare pauses """
		return re.findall(r"[\w']+", text) #text.replace("\n", "").split(",")

class TextExtractor:
	def __init__(self):
		pass
		
	def __call__(self,inputFile):
		# TODO: Better seperate headers and columns, especially on the first page
		# TODO: Better recognize when the citations are going to occur and cut them out
		pdfFileObj = open(inputFile + '.pdf','rb')	 #'rb' for read binary mode
		pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
		extractedText = """"""
		for page in range(0,pdfReader.numPages):
			extractedText += pdfReader.getPage(page).extractText()
		return extractedText

def parser():
	parser = argparse.ArgumentParser()
	parser.add_argument('-f', '--file',required=True)
	parser.add_argument('-r', '--engineRate', nargs = '?', type=int,default = -2)
	parser.add_argument('-p', '--pause', nargs = '?', type=int, default = 150)
	args = parser.parse_args()

	input_file = args.file
	output_file = input_file
	pdfExtractor = TextExtractor()
	audioConverter = SpeechToWav(args.pause,args.engineRate)
	audioConverter(text = pdfExtractor(inputFile = input_file),outputFile = output_file)
	
if __name__ == '__main__':
	parser()
	