from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_HtmlFile.html"
outputFile = "HtmlToPdf.pdf"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile, FileFormat.Html, XHTMLValidationType.none)
#Save to file.
document.SaveToFile(outputFile, FileFormat.PDF)
document.Close()


inputFile = "./Data/Template_HtmlFile.html"
outputFile =  "HtmlToPdf_PS.pdf"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile, FileFormat.Html, XHTMLValidationType.none)
#Save to file.
Ps = ToPdfParameterList()
Ps.UsePSCoversion = True
document.SaveToFile(outputFile, Ps)
document.Close()
