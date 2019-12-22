from comtypes.client import CreateObject
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import time
import xml.etree.ElementTree as ET
import re
import argparse
import os
import io
import sys
from nltk.tokenize import sent_tokenize

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
				phonemeDict[tok[0].lower()] = tok[1].lower().rstrip(',') # Dirty hack
		return phonemeDict
	
	def __call__(self,text):
		SSML_document = self._header()
		for utterance in text:
			parent_tag = self._pronounce(utterance,SSML_document)
			SSML_document.append(parent_tag)
		ET.dump(SSML_document)
		return SSML_document
		
	def _pause(self,parent_tag, pause_time):
		return 	ET.fromstring("<break time=\"" + pause_time + "ms\" />") # ET.SubElement(parent_tag,"break",{"time":str(self.pause)+"ms"})

	def _header(self):
		return ET.Element("speak",{"version":"1.0", "xmlns":"http://www.w3.org/2001/10/synthesis", "xml:lang":"en-US"})
	
	# TODO: Add rate https://docs.microsoft.com/en-us/cortana/skills/speech-synthesis-markup-language#prosody-element
	def _rate(self):
		pass
	
	# TODO: Add pitch 
	def _pitch(self):
		pass
	
	# I think there is something wrong with the way words are being tacked on
	def _pronounce(self,word,parent_tag):
		if word in self.phonemeDict:
			#print(word, self.phonemeDict[word])
			return ET.fromstring('<phoneme alphabet="ipa" ph=\'' + self.phonemeDict[word] + '\'> </phoneme>')#ET.SubElement(parent_tag,"phoneme",{"alphabet":"ipa","ph":self.phonemeDict[word]})#<phoneme alphabet="string" ph="string"></phoneme>
		else:
			return ET.fromstring('<s>' + word +'</s>')
	# Nice to have: Transform acronyms into their pronunciation (See say as tag)

class TextProcesser:
	def __init__(self,type):
		self.type = type
		self.preprocessText = self._preprocessText
		self.tokenizeText = self._wordTokenizer
		if type == "sent":
			self.tokenizeText = sent_tokenize
		elif type == "word":
			self.tokenizeText = self._wordTokenizer
	
	def _wrapAround(self, text):
		return text.replace('-\n','')
		
	def _removeNewlines(self, text):
		return text.replace('\n',' ')

	def _wordTokenizer(self, text):
		return re.findall(r"[\w']+", text)
	
	def _preprocessText(self, text):
		text = self._wrapAround(text)
		return text.lower()
	
	def __call__(self,text):
		text = self.preprocessText(text)
		return self.tokenizeText(text)

class TTS:
	def __init__(self,pause,engineRate,phonemeFile):
		self.engine = CreateObject("SAPI.SpVoice")
		self.stream = CreateObject("SAPI.SpFileStream")
		self.SSML_generator = SSML_Generator(pause,phonemeFile)
		self.textProcesser = TextProcesser("word")
		# This import has to be done in a specific order, otherwise Windows is not happy
		from comtypes.gen import SpeechLib
		# Because of that, below is a workaround to use SpeechLib throughout the object.
		self.SpeechLib = SpeechLib
		self.engine.Rate = engineRate
	
	# Add option to change voice
	def change_voice(self,voice_id):
		pass
	
	def __call__(self,text,outputFile):
		# https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms723606(v%3Dvs.85)
		self.stream.Open(outputFile + ".wav", self.SpeechLib.SSFMCreateForWrite)
		self.engine.AudioOutputStream = self.stream
		import pdb;pdb.set_trace()
		text = self.textProcesser(text)
		text = self.SSML_generator(text)
		text = ET.tostring(text,encoding='utf8', method='xml').decode('utf-8')
		self.engine.speak(text)
		self.stream.Close()

class TextExtractor:
	def __init__(self):
		pass
		
	def __call__(self, path):
		rsrcmgr = PDFResourceManager()
		retstr = StringIO()
		codec = 'utf-8'
		laparams = LAParams()
		device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
		fp = open(path, 'rb')
		interpreter = PDFPageInterpreter(rsrcmgr, device)
		password = ""
		maxpages = 0
		caching = True
		pagenos=set()

		for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
			interpreter.process_page(page)

		text = retstr.getvalue()

		fp.close()
		device.close()
		retstr.close()
		return text

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
	audioConverter = TTS(args.pause,args.engineRate,args.phonemeFile)
	audioConverter(text = pdfExtractor(input_file),outputFile = output_file)
	
if __name__ == '__main__':
	parser()
	'''
	#input_file = args.file
	output_file = "hwet"
	#pdfExtractor = TextExtractor()
	cwd = os.getcwd()
	phonetic_lib_dir = cwd #os.path.join(cwd,"cmudict-ipa")
	audioConverter = SpeechToWav(150,-2,os.path.join(phonetic_lib_dir,"cmudict-0.7b-ipa.txt"))
	audioConverter(text = "Hello world, here is an extended test",outputFile = output_file)
	'''
	