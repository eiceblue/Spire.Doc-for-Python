from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Sample.docx"
outputFile = "FindAndHighlight.docx"

#Create word document
document = Document()

#Load the document from disk.
document.LoadFromFile(inputFile)

#Find text
textSelections = document.FindAllString("word", False, True)

#Set hightlight
for selection in textSelections:
    selection.GetAsOneRange().CharacterFormat.HighlightColor = Color.get_Yellow()

#Save doc file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()        
