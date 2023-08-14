from spire.doc import *
from spire.doc.common import *

outputFile = "InsertPageRefField.docx"
inputFile = "./Data/PageRef.docx"

# Open a Word document
document = Document()
document.LoadFromFile(inputFile)

# Get the first section
section = document.LastSection

par = section.AddParagraph()

# Add page ref field
field = par.AppendField("pageRef", FieldType.FieldPageRef)

# Set field code
field.Code = "PAGEREF  bookmark1 \\# \"0\" \\* Arabic  \\* MERGEFORMAT"

# Update field
document.IsUpdateFields = True

document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
