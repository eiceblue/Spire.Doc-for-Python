from spire.doc import *
from spire.doc.common import *

outputFile = "FormFieldsProperties.docx"
inputFile = "./Data/FillFormField.doc"

# Open a Word document
document = Document()
document.LoadFromFile(inputFile)

# Get the first section
section = document.Sections[0]

# Get FormField by index
formField = section.Body.FormFields[1]

if formField.Type == FieldType.FieldFormTextInput:
    formField.Text = "My name is " + formField.Name
    formField.CharacterFormat.TextColor = Color.get_Red()
    formField.CharacterFormat.Italic = True

document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
