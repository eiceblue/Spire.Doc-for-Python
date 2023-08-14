
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/InputHtmlFile.html"
outputFile = "HtmlFileToWord.docx"

#Open an html file.
document = Document()
document.LoadFromFile(inputFile, FileFormat.Html, XHTMLValidationType.none)
#Save it to a Word document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()

