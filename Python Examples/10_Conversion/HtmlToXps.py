from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Template_HtmlFile.html"
outputFile = "HtmlToXps.xps"
#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile, FileFormat.Html, XHTMLValidationType.none)
#Save to file.
document.SaveToFile(outputFile, FileFormat.XPS)
document.Close()
