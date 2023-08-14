from spire.doc import *
from spire.doc.common import *

inputFile =  "./Data/Template_OdtFile.odt"
outputFile = "OdtToWord.docx"
#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Save to Docx file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
