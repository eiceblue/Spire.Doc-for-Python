from spire.doc import *
from spire.doc.common import *


outputFile = "SetColumnWidth.docx"
inputFile = "./Data/TableSample.docx"

# Create a document and load file
document = Document()
document.LoadFromFile(inputFile)

section = document.Sections[0]
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

# Traverse the first column
for i in range(table.Rows.Count):
    # Set the width and type of the cell
    table.Rows[i].Cells[0].SetCellWidth(200, CellWidthType.Point)

# Save to file
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
