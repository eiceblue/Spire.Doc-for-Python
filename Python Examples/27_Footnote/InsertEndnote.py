from spire.doc import *
from spire.doc.common import *

outputFile = "InsertEndnote.docx"
inputFile = "./Data/InsertEndnote.doc"

#Create a document and load file
doc = Document()
doc.LoadFromFile(inputFile)
s = doc.Sections[0]
p = s.Paragraphs[1]

#add endnote
endnote = p.AppendFootnote(FootnoteType.Endnote)

#append text
text = endnote.TextBody.AddParagraph().AppendText("Reference: Wikipedia")

#set text format
text.CharacterFormat.FontName = "Impact"
text.CharacterFormat.FontSize = 14
text.CharacterFormat.TextColor = Color.get_DarkOrange()

#Set marker format of endnote
endnote.MarkerCharacterFormat.FontName = "Calibri"
endnote.MarkerCharacterFormat.FontSize = 25
endnote.MarkerCharacterFormat.TextColor = Color.get_DarkBlue()

#Save the document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()

