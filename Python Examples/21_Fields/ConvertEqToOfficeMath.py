##有问题 documentObject.FieldText获取不到

from spire.doc import *
from spire.doc.common import *

inputFile = "Data/EQ.docx"
outputFile = "ConvertEqToOfficeMath.docx"

# Create a new document
document = Document()

# Load the document from a file
document.LoadFromFile(inputFile)

# Get the first paragraph of the first section in the document
paragraph = document.Sections.get_Item(0).Paragraphs.get_Item(0)

# Iterate through the child objects of the paragraph
i = 0
while i < paragraph.ChildObjects.Count:
    # Get the current document object
    documentObject = paragraph.ChildObjects[i]

    # Check if the document object is a field of type Equation
    if isinstance(documentObject,
                  Field) and documentObject.Type == FieldType.FieldEquation:
        # Convert the field to an OfficeMath object
        officeMath = OfficeMath.FromEqField(documentObject)

        # If conversion is successful, replace the field with the OfficeMath object
        if officeMath is not None:
            paragraph.ChildObjects.Remove(documentObject)
            paragraph.ChildObjects.Insert(i, officeMath)
    i += 1

# Save the modified document to a new file
document.SaveToFile(outputFile, FileFormat.Docx)

# Dispose of the document object
document.Dispose()
