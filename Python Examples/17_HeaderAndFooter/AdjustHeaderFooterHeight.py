from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/HeaderAndFooter.docx"
outputFile = "AdjustHeaderFooterHeight.docx"

# Load the document
doc = Document()
doc.LoadFromFile(inputFile)

# Get the first section
section = doc.Sections[0]

# Adjust the height of headers in the section
section.PageSetup.HeaderDistance = 100

# Adjust the height of footers in the section
section.PageSetup.FooterDistance = 100

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
