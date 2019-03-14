from comtypes.client import CreateObject
import PyPDF2
import time

class SpeechToWav:
	def __init__(self):
		self.engine = CreateObject("SAPI.SpVoice")
		self.stream = CreateObject("SAPI.SpFileStream")
		# This import has to be done in a specific order, otherwise Windows is not happy
		from comtypes.gen import SpeechLib
		# Because of that, below is a work around to use its functions throughout the object.
		self.SpeechLib = SpeechLib
		self.engine.Rate = -1

		
	# TODO: Slow down the rate even more
	# TODO: Add way to easily change the voice
	# TODO: Add pauses between commas
	def __call__(self,text,outputFile):
		# https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms723606(v%3Dvs.85)
		self.stream.Open(outputFile + ".wav", self.SpeechLib.SSFMCreateForWrite)
		self.engine.AudioOutputStream = self.stream
		for utterance in text:
			self.engine.speak(utterance)
		self.stream.Close()

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
		return self._post_processing(extractedText)
	
	def _post_processing(self,text):
		""" Used to remove side effects from columns, the new lines add bizzare pauses """
		return text.replace("\n", "").replace(" ", ",").split(",")
	
	def _delimit_by_comma(self,text):
		return text.split(",")
	def _new_line_removal(self,text):
		return text.replace("\n", "")

	# TODO: Add a conversion that replaces numbers with corresponding strings (Namely thousands, it struggles with those)
	# TODO: Recognize and replace references
	# TODO: Figure out when it says "dot" instead of just using it as a period and remove those
	# Nice to have: Transform acronyms into their pronounciation
	
# TODO: Add an easy way to interface with files (CLI?)
input_file = "test"
output_file = input_file
pdfExtractor = TextExtractor()
audioConverter = SpeechToWav()
audioConverter(text = pdfExtractor(inputFile = input_file),outputFile = output_file)


