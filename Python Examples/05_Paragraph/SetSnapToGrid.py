import unittest
from spire.doc import *
from spire.doc.common import *

outputFile = "SetSnapToGrid.docx"

# Create a new instance of the Document class.
doc = Document()

# Add a new section to the document.
section = doc.AddSection()

# Set the grid type of the page setup in the section to "LinesOnly".
section.PageSetup.GridType = GridPitchType.LinesOnly

# Set the number of lines per page in the section to 15.
section.PageSetup.LinesPerPage = 15

# Add a new paragraph to the section.
paragraph = section.AddParagraph()

# Append text to the paragraph.
paragraph.AppendText(
    "With Spire.Doc, you can generate, modify, convert, render and print documents without utilizing Microsoft Word®. But you need MS Word viewer to view the resultant document. "
)

# Set the "SnapToGrid" property of the paragraph's format to true.
paragraph.Format.SnapToGrid = True

# Save the document to a file with the specified file name and format (Docx2013).
doc.SaveToFile(outputFile, FileFormat.Docx2013)

# Clean up resources used by the document.
doc.Dispose()
