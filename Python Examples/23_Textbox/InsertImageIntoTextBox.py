from spire.doc import *
from spire.doc.common import *

outputFile = "InsertImageIntoTextBox.docx"

#Create a new document
doc = Document()
section = doc.AddSection()
paragraph = section.AddParagraph()

#Append a textbox to paragraph
tb = paragraph.AppendTextBox(220, 220)

#Set the position of the textbox
tb.Format.HorizontalOrigin = HorizontalOrigin.Page
tb.Format.HorizontalPosition = 50
tb.Format.VerticalOrigin = VerticalOrigin.Page
tb.Format.VerticalPosition = 50

#Set the fill effect of textbox as picture
tb.Format.FillEfects.Type = BackgroundType.Picture

#Fill the textbox with a picture
tb.Format.FillEfects.SetPicture("./Data/Spire.Doc.png")

#Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
