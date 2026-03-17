from spire.doc import *
from spire.doc.common import *

# Create a new Document instance
document = Document()

# Load a Word document from file
document.LoadFromFile("Data/Sample.docx")

# Note: This default font is used only when the font specified in the document is not found in the font cache.
document.DefaultSubstitutionFontName = "Arial"

# Save the document as a PDF file
document.SaveToFile("Sample.pdf", FileFormat.PDF)

# Close the document to release resources
document.Close()

# Dispose of the document object to free memory
document.Dispose()