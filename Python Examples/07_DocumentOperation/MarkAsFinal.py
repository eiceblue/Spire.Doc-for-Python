from spire.doc import *
from spire.doc.common import *

# Create a new Document object
document = Document()

# Load the document from the input file
document.LoadFromFile("Data/Insert.docx")

# Get the custom document properties
customProperties = document.CustomDocumentProperties

# Add _MarkAsFinal custom document property
customProperties.Add("_MarkAsFinal", Boolean(True))

# Save the document to the output file as PDF
document.SaveToFile("MarkAsFinal.docx", FileFormat.Docx2013)

# Close and dispose the document object
document.Close()
document.Dispose()