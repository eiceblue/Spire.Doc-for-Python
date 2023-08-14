from spire.doc import *
from spire.doc.common import *

inputFile =  "./Data/Template_TxtFile.txt"
outputFile = "TxtToWord.docx"
        
#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Save the file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()