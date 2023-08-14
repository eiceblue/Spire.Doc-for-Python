from spire.doc import *
from spire.doc.common import *

outputFile = "TextBoxFormat.docx"

#Create a new document
doc = Document()
sec = doc.AddSection()

#Add a text box and append sample text
TB = doc.Sections[0].AddParagraph().AppendTextBox(310, 90)
para = TB.Body.AddParagraph()
TR = para.AppendText("Using Spire.Doc, developers will find " + "a simple and effective method to endow their applications with rich MS Word features. ")
TR.CharacterFormat.FontName = "Cambria "
TR.CharacterFormat.FontSize = 13

#Set exact position for the text box
TB.Format.HorizontalOrigin = HorizontalOrigin.Page
TB.Format.HorizontalPosition = 120
TB.Format.VerticalOrigin = VerticalOrigin.Page
TB.Format.VerticalPosition = 100

#Set line style for the text box
TB.Format.LineStyle = TextBoxLineStyle.Double
TB.Format.LineColor = Color.get_CornflowerBlue()
TB.Format.LineDashing = LineDashing.Solid
TB.Format.LineWidth = 5

#Set internal margin for the text box
TB.Format.InternalMargin.Top = 15
TB.Format.InternalMargin.Bottom = 10
TB.Format.InternalMargin.Left = 12
TB.Format.InternalMargin.Right = 10

#Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()