from comtypes.client import CreateObject
import PyPDF2
import time
import xml.etree.ElementTree as ET
import re
import argparse
import os
import io
import sys
# TODO: Make it cross system by removing dependency on comtypes
# (https://stackoverflow.com/questions/39014468/how-to-save-the-output-of-pyttsx-to-wav-file)
# https://pypi.org/project/pyttsx3/ (looks unsupported..)

# https://www.w3.org/TR/speech-synthesis/
# TODO: Add way to easily modify the voice 
class SSML_Generator:
	def __init__(self,pause,phonemeFile):
		self.pause = pause
		if isinstance(phonemeFile,str):
			print("Loading dictionary")
			self.phonemeDict = self._load_phonemes(phonemeFile)
			print(len(self.phonemeDict))
		else:
			self.phonemeDict = {}
		
	def _load_phonemes(self, phonemeFile):
		phonemeDict = {}
		with io.open(phonemeFile, 'r',encoding='utf-8') as f:
			for line in f:
				tok = line.split()
				#print(len(tok))
				phonemeDict[tok[0].lower()] = tok[1].lower()
		return phonemeDict
	
	def __call__(self,text):
		SSML_document = self._header()
		for utterance in text:
			parent_tag = self._pronounce(utterance,SSML_document)
			#parent_tag.tail = self._pause(parent_tag)
			SSML_document.append(parent_tag)
		ET.dump(SSML_document)
		return SSML_document
		
	def _pause(self,parent_tag):
		return 	ET.fromstring("<break time=\"150ms\" />") # ET.SubElement(parent_tag,"break",{"time":str(self.pause)+"ms"})

	def _header(self):
		return ET.Element("speak",{"version":"1.0", "xmlns":"http://www.w3.org/2001/10/synthesis", "xml:lang":"en-US"})
	
	# TODO: Add rate https://docs.microsoft.com/en-us/cortana/skills/speech-synthesis-markup-language#prosody-element
	def _rate(self):
		pass
	
	# TODO: Add pitch 
	def _pitch(self):
		pass
	
	def _pronounce(self,word,parent_tag):
		if word in self.phonemeDict:
			print(self.phonemeDict[word])
			return ET.fromstring("<phoneme alphabet=\"ipa\" ph=\"" + self.phonemeDict[word] + "\"> </phoneme>")#ET.SubElement(parent_tag,"phoneme",{"alphabet":"ipa","ph":self.phonemeDict[word]})#<phoneme alphabet="string" ph="string"></phoneme>
		else:
			return parent_tag
	# Nice to have: Transform acronyms into their pronunciation (See say as tag)
	
class SpeechToWav:
	def __init__(self,pause,engineRate,phonemeFile):
		self.engine = CreateObject("SAPI.SpVoice")
		self.stream = CreateObject("SAPI.SpFileStream")
		self.SSML_generator = SSML_Generator(pause,phonemeFile)
		# This import has to be done in a specific order, otherwise Windows is not happy
		from comtypes.gen import SpeechLib
		# Because of that, below is a work around to use its functions throughout the object.
		self.SpeechLib = SpeechLib
		self.engine.Rate = engineRate
	
	# Add option to change voice
	def change_voice(self,voice_id):
		pass
	
	def __call__(self,text,outputFile):
		# https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms723606(v%3Dvs.85)
		self.stream.Open(outputFile + ".wav", self.SpeechLib.SSFMCreateForWrite)
		self.engine.AudioOutputStream = self.stream
		text = self._text_processing(text)
		text = self.SSML_generator(text)
		text = ET.tostring(text,encoding='utf8', method='xml').decode('utf-8')
		self.engine.speak(text)
		self.stream.Close()
	
	
	def _text_processing(self,text):
		""" Used to remove side effects from columns, the new lines add bizzare pauses """
		return re.findall(r"[\w']+", text.lower())

class TextExtractor:
	def __init__(self):
		pass
		
	def __call__(self,inputFile):
		# TODO: Find good ways to identify the layout and different sections
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
	parser.add_argument('-i', '--phonemeFile',nargs = '?',default = None)
	parser.add_argument('-r', '--engineRate', nargs = '?', type=int,default = -2)
	parser.add_argument('-p', '--pause', nargs = '?', type=int, default = 150)
	args = parser.parse_args()

	input_file = args.file
	output_file = input_file
	pdfExtractor = TextExtractor()
	audioConverter = SpeechToWav(args.pause,args.engineRate,args.phonemeFile)
	audioConverter(text = pdfExtractor(inputFile = input_file),outputFile = output_file)
	
if __name__ == '__main__':
	#parser()
	#input_file = args.file
	output_file = "hwet"
	#pdfExtractor = TextExtractor()
	cwd = os.getcwd()
	phonetic_lib_dir = os.path.join(cwd,"cmudict-ipa")
	audioConverter = SpeechToWav(150,-2,os.path.join(phonetic_lib_dir,"cmudict-0.7b-ipa.txt"))
	audioConverter(text = "Hello world extended test",outputFile = output_file)
	