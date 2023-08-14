from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Sample_UTF-7.txt"
outputFile = "LoadTextWithEncoding.docx"

#Create word document
document = Document()
#Load the text file 
document.LoadText(inputFile, Encoding.get_UTF7())
#Save the document
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()