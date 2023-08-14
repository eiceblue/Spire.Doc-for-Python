from spire.doc import *
from spire.doc.common import *


outputFile = "TextWaterMark.docx"
inputFile = "./Data/Template.docx"

#Open a Word document as template.
document = Document()
document.LoadFromFile(inputFile)

#Insert text watermark.
txtWatermark = TextWatermark()
txtWatermark.Text = "E-iceblue"
txtWatermark.FontSize = 95
txtWatermark.Color = Color.get_Blue()
txtWatermark.Layout = WatermarkLayout.Diagonal
document.Watermark = txtWatermark

#Save as docx file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()

        
