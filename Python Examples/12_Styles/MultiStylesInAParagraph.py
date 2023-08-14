from spire.doc import *
from spire.doc.common import *

outputFile = "MultiStylesInAParagraph.docx"
#Create a Word document
doc = Document()

#Add a section
section = doc.AddSection()

#Add a paragraph
para = section.AddParagraph()

#Add a text range 1 and set its style
range = para.AppendText("Spire.Doc for .NET ")
range.CharacterFormat.FontName = "Calibri"
range.CharacterFormat.FontSize = 16
range.CharacterFormat.TextColor = Color.get_Blue()
range.CharacterFormat.Bold = True
range.CharacterFormat.UnderlineStyle = UnderlineStyle.Single

#Add a text range 2 and set its style
range = para.AppendText("is a professional Word .NET library")
range.CharacterFormat.FontName = "Calibri"
range.CharacterFormat.FontSize = 15

#Save the Word document
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()

