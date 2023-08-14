
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_1.docx"
outputFile =  "WordToPDFA.pdf"
        
#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Set the Conformance-level of the Pdf file to PDF_A1B.
toPdf = ToPdfParameterList()
toPdf.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B
#Save the file.
document.SaveToFile(outputFile, toPdf)
document.Close()

