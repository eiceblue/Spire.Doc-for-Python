from spire.doc import *
from spire.doc.common import *

outputFile = "InsertNoneField.docx"
inputFile = "./Data/SampleB_2.docx"

# Open a Word document.
document = Document()
document.LoadFromFile(inputFile)

# Get the first section
section = document.Sections[0]

par = section.AddParagraph()

# Add a none field
field = par.AppendField('', FieldType.FieldNone)

document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
