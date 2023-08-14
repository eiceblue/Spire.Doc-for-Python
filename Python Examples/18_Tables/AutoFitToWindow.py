
from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/TableSample.docx"
outputFile = "AutoFitToWindow.docx"

# Create a document
document = Document()

# Load file
document.LoadFromFile(inputFile)
section = document.Sections[0]
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

# Automatically fit the table to the active window width
table.AutoFit(AutoFitBehaviorType.AutoFitToWindow)

# Save to file and launch it
document.SaveToFile(outputFile)
document.Close()
