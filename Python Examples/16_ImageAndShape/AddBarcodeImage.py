from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/SampleB_2.docx"
outputFile = "AddBarcodeImage.docx"

#Open a Word document
document = Document()
document.LoadFromFile(inputFile)

imgPath = "./Data/barcode.png"

#Add barcode image
picture = document.Sections[0].AddParagraph().AppendPicture(imgPath)

#Save to file
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()

