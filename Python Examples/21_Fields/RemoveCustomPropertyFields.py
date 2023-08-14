from spire.doc import *
from spire.doc.common import *

outputFile = "RemoveCustomPropertyFields.docx"
inputFile = "./Data/RemoveCustomPropertyFields.docx"

# Create Word document.
document = Document()

# Load the file from disk.
document.LoadFromFile(inputFile)

# Get custom document properties object.
cdp = document.CustomDocumentProperties

# Remove all custom property fields in the document.
i = 0
while i < cdp.Count:
    cdp.Remove(cdp[i].Name)

document.IsUpdateFields = True

# Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
