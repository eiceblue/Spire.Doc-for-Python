from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Template_Docx_1.docx"
outputFile = "SetSpacing.docx"

#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)
#Add the text strings to the paragraph and set the style.
para = Paragraph(document)
textRange1 = para.AppendText("This is an inserted paragraph.")
textRange1.CharacterFormat.TextColor = Color.get_Blue()
textRange1.CharacterFormat.FontSize = 15

#set the spacing before and after.
para.Format.BeforeAutoSpacing = False
para.Format.BeforeSpacing = 10
para.Format.AfterAutoSpacing = False
para.Format.AfterSpacing = 10

#insert the added paragraph to the word document.
document.Sections[0].Paragraphs.Insert(1, para)

#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()

