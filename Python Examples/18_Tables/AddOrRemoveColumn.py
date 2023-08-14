from spire.doc import *
from spire.doc.common import *

def AddColumn(table, columnIndex):
    for r in range(table.Rows.Count):
        addCell = TableCell(table.Document)
        table.Rows[r].Cells.Insert(columnIndex, addCell)
def RemoveColumn(table, columnIndex):
    for r in range(table.Rows.Count):
        table.Rows[r].Cells.RemoveAt(columnIndex)


inputFile = "./Data/Template_N2.docx"
outputFile = "AddOrRemoveColumn.docx"

# Load the document from disk.
doc = Document()
doc.LoadFromFile(inputFile)

# Access the first section
section = doc.Sections[0]

# Access the first table
table = section.Tables[0] if isinstance(
    section.Tables[0], Table) else None

# Add a blank column
columnIndex1 = 0
AddColumn(table, columnIndex1)

# Remove a column
columnIndex2 = 2
RemoveColumn(table, columnIndex2)

# Save the Word file
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()




