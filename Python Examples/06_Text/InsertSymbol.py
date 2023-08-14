from spire.doc import *
from spire.doc.common import *

outputFile = "InsertSymbol.docx"

#Create Word document.
document = Document()
#Add a section.
section = document.AddSection()
#Add a paragraph.
paragraph = section.AddParagraph()
#Use unicode characters to create symbol Ä.
tr = paragraph.AppendText(str('\u00c4'))
#Set the color of symbol Ä.
tr.CharacterFormat.TextColor = Color.get_Red()
#Add symbol Ë.
paragraph.AppendText(str('\u00cb'))
#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()