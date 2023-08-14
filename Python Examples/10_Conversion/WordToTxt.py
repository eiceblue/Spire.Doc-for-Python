from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_1.docx"
outputFile =  "WordToTxt.txt"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Save the file.
document.SaveToFile(outputFile, FileFormat.Txt)
document.Close()
