from comtypes.client import CreateObject
import PyPDF2
import time
import xml.etree.ElementTree as ET
import re

# https://www.w3.org/TR/speech-synthesis/
# https://stackoverflow.com/questions/14698227/is-there-some-character-i-can-insert-to-get-a-longer-pause-with-sapi5-tts
# TODO: Add way to easily modify the voice 
class SSML_Generator:
	def __init__(self,pause = 150):
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
	def __init__(self):
		self.engine = CreateObject("SAPI.SpVoice")
		self.stream = CreateObject("SAPI.SpFileStream")
		self.SSML_generator = SSML_Generator()
		# This import has to be done in a specific order, otherwise Windows is not happy
		from comtypes.gen import SpeechLib
		# Because of that, below is a work around to use its functions throughout the object.
		self.SpeechLib = SpeechLib
		self.engine.Rate = -2
		self.pause_time_in_ms = 30
		
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
		pdfFileObj = open(inputFile + '.pdf','rb')     #'rb' for read binary mode
		pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
		extractedText = """"""
		for page in range(0,pdfReader.numPages):
			extractedText += pdfReader.getPage(page).extractText()
		return extractedText
	

# TODO: Add an easy way to interface with files (CLI?)
# https://blog.sicara.com/perfect-python-command-line-interfaces-7d5d4efad6a2

input_file = "test"
output_file = input_file
pdfExtractor = TextExtractor()
audioConverter = SpeechToWav()
audioConverter(text = pdfExtractor(inputFile = input_file),outputFile = output_file)
"""

fle = "hw"
sg = SSML_Generator()
a = sg(["Hello","World","Extended","Test"])
ac = SpeechToWav()
#ET.dump(a)
ac("Hello world, here is an extended test.\n This is the second line",fle)
"""

