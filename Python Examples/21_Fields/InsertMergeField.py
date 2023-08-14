
from spire.doc import *
from spire.doc.common import *

outputFile = "InsertMergeField.docx"
inputFile = "./Data/SampleB_2.docx"

# Open a Word document
document = Document()
document.LoadFromFile(inputFile)

# Get the first section
section = document.Sections[0]

par = section.AddParagraph()

# Add merge field in the paragraph
field = MergeField(par.AppendField(
    "MyFieldName", FieldType.FieldMergeField))

# Save to file
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
