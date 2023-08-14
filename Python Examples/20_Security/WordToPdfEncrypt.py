from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Template_Docx_2.docx"
outputFile = "WordToPdfEncrypt.pdf"

#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)

#Create an instance of ToPdfParameterList.
toPdf = ToPdfParameterList()

#Set the user password for the resulted PDF file.
toPdf.PdfSecurity.Encrypt("e-iceblue")

#Save to file.
document.SaveToFile(outputFile, toPdf)
document.Close()
