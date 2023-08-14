from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_5.docx"
outputFile =  "KeepHiddenText_PS.pdf"
#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#When convert to PDF file, set the property IsHidden as true.
pdf = ToPdfParameterList()
pdf.IsHidden = True
#Save to file.
document.SaveToFile(outputFile, pdf)
document.Close()

