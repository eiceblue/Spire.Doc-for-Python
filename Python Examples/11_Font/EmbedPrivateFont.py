from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/BlankTemplate.docx"
outputFile = "EmbedPrivateFont.docx"

#Load the document
doc = Document()
doc.LoadFromFile(inputFile)

#Get the first section and add a paragraph
section = doc.Sections[0]
p = section.AddParagraph()

#Append text to the paragraph, then set the font name and font size
txtRange = p.AppendText("Spire.Doc for Python is a professional Word Python API specifically designed for developers to create, read, write, convert, and compare Word documents with fast and high-quality performance.")
txtRange.CharacterFormat.FontName = "PT Serif Caption"
txtRange.CharacterFormat.FontSize = 20

#Allow embedding font in document
doc.EmbedFontsInFile = True

#Embed private font from font file into the document
doc.PrivateFontList.append(PrivateFontPath("PT Serif Caption", "./Data/PT Serif Caption.ttf"))

#Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()