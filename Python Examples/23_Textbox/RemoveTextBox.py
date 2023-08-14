from spire.doc import *
from spire.doc.common import *

outputFile = "RemoveTextBox.docx"
inputFile = "./Data/TextBoxTemplate.docx"

#Load the document
doc = Document()
doc.LoadFromFile(inputFile)

#Remove the first text box
doc.TextBoxes.RemoveAt(0)

#Clear all the text boxes
#Doc.TextBoxes.Clear()
#Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
