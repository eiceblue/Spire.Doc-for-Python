from spire.doc import *
from spire.doc.common import *

outputFile = "ConvertIfFieldToText.docx"
inputFile = "./Data/IfFieldSample.docx"

# Open a Word document
document = Document()
document.LoadFromFile(inputFile)

# Get all fields in document
fields = document.Fields

for i in range(fields.Count):
    field = fields[i]
    if field.Type == FieldType.FieldIf:
        original = field if isinstance(field, TextRange) else None
        # Get field text
        text = field.FieldText
        # Create a new textRange and set its format
        textRange = TextRange(document)
        textRange.Text = text
        textRange.CharacterFormat.FontName = original.CharacterFormat.FontName
        textRange.CharacterFormat.FontSize = original.CharacterFormat.FontSize

        par = field.OwnerParagraph
        # Get the index of the if field
        index = par.ChildObjects.IndexOf(field)
        # Remove if field via index
        par.ChildObjects.RemoveAt(index)
        # Insert field text at the position of if field
        par.ChildObjects.Insert(index, textRange)
# Save doc file
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
