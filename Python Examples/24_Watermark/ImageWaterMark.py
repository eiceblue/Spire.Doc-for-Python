from spire.doc import *
from spire.doc.common import *

outputFile = "ImageWaterMark.docx"
inputFile = "./Data/Template.docx"

#Open a Word document as template.
document = Document()
document.LoadFromFile(inputFile)

#Insert the imgae watermark.
picture = PictureWatermark()
picture.SetPicture("./Data/ImageWatermark.png")
picture.Scaling = 250
picture.IsWashout = False
document.Watermark = picture

#Save as docx file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
        
