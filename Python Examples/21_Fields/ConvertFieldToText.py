from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Fields.docx"
outputFile = "ConvertFieldToText.docx"

# Load word document
document = Document()
document.LoadFromFile(inputFile)

# Get all fields in document
fields = document.Fields
count = fields.Count

for i in range(0, count):
    field = fields[0]
    s = field.FieldText
    index = field.OwnerParagraph.ChildObjects.IndexOf(field)
    textRange = TextRange(document)
    textRange.Text = s
    textRange.CharacterFormat.FontSize = 24

    field.OwnerParagraph.ChildObjects.Insert(index, textRange)
    field.OwnerParagraph.ChildObjects.Remove(field)

# Save doc file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
