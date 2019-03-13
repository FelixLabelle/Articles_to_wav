from comtypes.client import CreateObject
import PyPDF2

class SpeechToWav:
	def __init__(self):
		self.engine = CreateObject("SAPI.SpVoice")
		self.stream = CreateObject("SAPI.SpFileStream")
		# https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms723606(v%3Dvs.85)

	def __call__(self,text,outputFile):
		from comtypes.gen import SpeechLib
		self.engine.Rate = 0
		self.stream.Open(outputFile + ".wav", SpeechLib.SSFMCreateForWrite)
		self.engine.AudioOutputStream = self.stream
		self.engine.speak(text)
		self.stream.Close()

class TextExtractor:
	def __init__(self):
		pass
	def __call__(self,inputFile):
		pdfFileObj = open(inputFile + '.pdf','rb')     #'rb' for read binary mode
		pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
		extractedText = """"""
		for page in range(0,pdfReader.numPages):
			extractedText += pdfReader.getPage(page).extractText()
		return self._post_processing(extractedText)
	
	def _post_processing(self,text):
		""" Used to remove side effects from columns, the new lines add bizzare pauses """
		return text.replace("\n", "")

input_file = "test"
output_file = input_file
pdfExtractor = TextExtractor()
audioConverter = SpeechToWav()
audioConverter(text = pdfExtractor(inputFile = input_file),outputFile = output_file)


