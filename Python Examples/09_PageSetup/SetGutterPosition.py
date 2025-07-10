from spire.doc import *
from spire.doc.common import *

inputFile = "data/Sample.docx"
outputFile = "SetGutterPosition.docx"

# Create a new Document object.
document = Document()

# Load a Word document from a specified file path.
document.LoadFromFile(inputFile)

# Get the first section of the document.
section = document.Sections[0]

# Set the top gutter option to true for the section's page setup.
section.PageSetup.IsTopGutter = True

# Set the width of the gutter in points (100f).
section.PageSetup.Gutter = 100

# Save the modified document to the specified output file path in DOCX format.
document.SaveToFile(outputFile, FileFormat.Docx)

# Dispose the existing document object.
document.Dispose()
