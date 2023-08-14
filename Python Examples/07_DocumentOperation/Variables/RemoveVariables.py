from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_6.docx"
outputFile = "RemoveVariables.docx"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Remove the variable by name.
document.Variables.Remove("A1")
document.IsUpdateFields = True
#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()