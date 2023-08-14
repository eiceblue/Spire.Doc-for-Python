from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Sample.docx"
outputFile = "InsertNewText.docx"

#Load Document   
doc = Document()
doc.LoadFromFile(inputFile)
#Find all the text string "New Zealand” from the sample document
selections = doc.FindAllString("Word", True, True)
index = 0
#Defines text range
trange = TextRange(doc)
#Insert new text string (New) after the searched text string
for selection in selections:
    trange = selection.GetAsOneRange()
    newrange = TextRange(doc)
    newrange.Text = ("(New text)")
    index = trange.OwnerParagraph.ChildObjects.IndexOf(trange)
    trange.OwnerParagraph.ChildObjects.Insert(index + 1, newrange)
#Find and highlight the newly added text string New
text = doc.FindAllString("New text", True, True)
for seletion in text:
    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.get_Yellow()
#Save the document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
