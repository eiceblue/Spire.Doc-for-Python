from spire.doc import *
from spire.doc.common import *

outputFile = "RemoveTable.docx"
inputFile = "./Data/Template.docx"

# Load the document
doc = Document()
doc.LoadFromFile(inputFile)

# Remove the first Table
doc.Sections[0].Tables.RemoveAt(0)

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
