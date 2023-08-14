
from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/ConvertedTemplate.docx"
outputFile = "ToXPS.xps"
#Create word document
document = Document()
document.LoadFromFile(inputFile)
#Save the document to a xps file.
document.SaveToFile(outputFile, FileFormat.XPS)
document.Close()
