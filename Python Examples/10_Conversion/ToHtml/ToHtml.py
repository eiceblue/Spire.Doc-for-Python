
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ToHtmlTemplate.docx"
outputFile = "ToHtml.html"
        
#Create word document
document = Document()
document.LoadFromFile(inputFile)
#Save doc file.
document.SaveToFile(outputFile, FileFormat.Html)
document.Close()

