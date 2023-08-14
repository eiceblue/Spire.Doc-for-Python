from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/AllowLatinTextWrapInMiddleOfAWord.docx"
outputFile = "AllowLatinTextWrapInMiddleOfAWord.docx"
       
document = Document()
document.LoadFromFile(inputFile)
para = document.Sections[0].Paragraphs[0]
#Allow Latin text to wrap in the middle of a word
para.Format.WordWrap = False
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
