from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/SampleB_2.docx"
outputFile = "AddImageToEachPage.docx"

#Open a Word document
document = Document()
document.LoadFromFile(inputFile)

imgPath = "./Data/Spire.Doc.png"

#Add a picture in footer and set it's position
picture = document.Sections[0].HeadersFooters.Footer.AddParagraph().AppendPicture(imgPath)

picture.VerticalOrigin = VerticalOrigin.Page
picture.HorizontalOrigin = HorizontalOrigin.Page
picture.VerticalAlignment = ShapeVerticalAlignment.Bottom
picture.TextWrappingStyle = TextWrappingStyle.none

#Add a textbox in footer and set it's positiion
textbox = document.Sections[0].HeadersFooters.Footer.AddParagraph().AppendTextBox(150, 20)
textbox.VerticalOrigin = VerticalOrigin.Page
textbox.HorizontalOrigin = HorizontalOrigin.Page
textbox.HorizontalPosition = 300
textbox.VerticalPosition = 700
textbox.Body.AddParagraph().AppendText("Welcome to E-iceblue")

#Save to file
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()

