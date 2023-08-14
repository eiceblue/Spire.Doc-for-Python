from spire.doc import *
from spire.doc.common import *


def WriteAllText(fpath: str, content: str):
    with open(fpath, 'w') as fp:
        fp.write(content)


outputFile = "GetFormFieldsCollection.txt"
inputFile = "./Data/FillFormField.doc"

sb = ''

# Open a Word document
document = Document()
document.LoadFromFile(inputFile)

# Get the first section
section = document.Sections[0]

formFields = section.Body.FormFields

sb = "The first section has " + str(formFields.Count) + " form fields."

WriteAllText(outputFile, sb)

document.Close()
