from spire.doc import *
from spire.doc.common import *

def WriteAllText(fpath: str, content: str):
    with open(fpath, 'w') as fp:
        fp.write(content)


outputFile = "GetFormFieldByName.txt"
inputFile = "./Data/FillFormField.doc"

sb = ''

# Open a Word document
document = Document()
document.LoadFromFile(inputFile)

# Get the first section
section = document.Sections[0]

# Get form field by name
formField = section.Body.FormFields["email"]

sb += "The name of the form field is " + formField.Name
sb += "\n"
sb += "The type of the form field is " + formField.FormFieldType.name
sb += "\n"

WriteAllText(outputFile, sb)
document.Close()
