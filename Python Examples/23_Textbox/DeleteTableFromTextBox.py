from spire.doc import *
from spire.doc.common import *

outputFile = "DeleteTableFromTextBox.docx"
inputFile = "./Data/TextBoxTable.docx"

#Load the document
doc = Document()
doc.LoadFromFile(inputFile)

#Get the first textbox
textbox = doc.TextBoxes[0]

#Remove the first table from the textbox
textbox.Body.Tables.RemoveAt(0)

#Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
