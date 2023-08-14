
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_5.docx"
outputFile = "DisableHyperlinks.pdf"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Create an instance of ToPdfParameterList.
pdf = ToPdfParameterList()
#Set DisableLink to true to remove the hyperlink effect for the result PDF page. 
#Set DisableLink to false to preserve the hyperlink effect for the result PDF page.
pdf.DisableLink = True
#Save to file.
document.SaveToFile(outputFile, pdf)
document.Close()

