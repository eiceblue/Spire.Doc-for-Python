
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ConvertedTemplate.docx"
outputFile = "ToPDF.pdf"
        
#Create word document
document = Document()
document.LoadFromFile(inputFile)
#Save the document to a PDF file.
document.SaveToFile(outputFile, FileFormat.PDF)
document.Close()