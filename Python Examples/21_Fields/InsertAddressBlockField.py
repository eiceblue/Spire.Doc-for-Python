from spire.doc import *
from spire.doc.common import *

outputFile = "InsertAddressBlockField.docx"
inputFile = "./Data/SampleB_2.docx"

# Open a Word document
document = Document()
document.LoadFromFile(inputFile)

# Get the first section
section = document.Sections[0]

par = section.AddParagraph()

# Add address block in the paragraph
field = par.AppendField("ADDRESSBLOCK", FieldType.FieldAddressBlock)

# Set field code
field.Code = "ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\""

# Save to file
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
