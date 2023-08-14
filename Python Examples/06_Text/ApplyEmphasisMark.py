from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Sample.docx"
outputFile = "ApplyEmphasisMark.docx"

# Create a new document and load from file
doc = Document()
doc.LoadFromFile(inputFile)
#Find text to emphasize
textSelections = doc.FindAllString("Spire.Doc for Python", False, True)

#Set emphasis mark to the found text
for selection in textSelections:
    selection.GetAsOneRange().CharacterFormat.EmphasisMark = Emphasis.Dot

#Save and launch the document
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()
