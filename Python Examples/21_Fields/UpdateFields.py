from spire.doc import *
from spire.doc.common import *

outputFile = "UpdateFields.docx"
inputFile = "./Data/IfFieldSample.docx"


#Open a Word document
document = Document()
document.LoadFromFile(inputFile)

#Update fields
document.IsUpdateFields = True

#Save doc file
document.SaveToFile(outputFile, FileFormat.Docx)
