from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_1.docx"
outputFile = "SetParagraphShading.docx"

#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)
#Get a paragraph.
paragaph = document.Sections[0].Paragraphs[0]

#Set background color for the paragraph.
paragaph.Format.BackColor = Color.get_Yellow()

#Set background color for the selected text of paragraph.
paragaph = document.Sections[0].Paragraphs[2]
selection = paragaph.Find("Christmas", True, False)
trange = selection.GetAsOneRange()
trange.CharacterFormat.TextBackgroundColor = Color.get_Yellow()

#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()