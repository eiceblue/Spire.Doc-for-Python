from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/TableSample.docx"
outputFile = "AutoFitToFixed.docx"

# Create a document
document = Document()
# Load file
document.LoadFromFile(inputFile)
section = document.Sections[0]
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None
# The table is set to a fixed size
table.AutoFit(AutoFitBehaviorType.FixedColumnWidths)
# Save to file
document.SaveToFile(outputFile)
document.Close()
