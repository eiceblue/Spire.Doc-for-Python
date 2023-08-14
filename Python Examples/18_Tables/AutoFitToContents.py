from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/TableSample.docx"
outputFile = "AutoFitToContents.docx"

# Create a document
document = Document()
# Load file
document.LoadFromFile(inputFile)

section = document.Sections[0]
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

# Automatically fit the table to the cell content
table.AutoFit(AutoFitBehaviorType.AutoFitToContents)

# Save to file and launch it
document.SaveToFile(outputFile)
document.Close()
