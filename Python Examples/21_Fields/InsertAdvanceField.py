from spire.doc import *
from spire.doc.common import *

outputFile = "InsertAdvanceField.docx"
inputFile = "./Data/SampleB_2.docx"

# Open a Word document.
document = Document()
document.LoadFromFile(inputFile)

# Get the first section
section = document.Sections[0]

par = section.AddParagraph()

# Add advance field
field = par.AppendField("Field", FieldType.FieldAdvance)

# Add field code
field.Code = "ADVANCE \\d 10 \\l 10 \\r 10 \\u 0 \\x 100 \\y 100 "

# Update field
document.IsUpdateFields = True

document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
