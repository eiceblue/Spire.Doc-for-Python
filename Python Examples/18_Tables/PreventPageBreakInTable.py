
from spire.doc import *
from spire.doc.common import *

outputFile = "PreventPageBreakInTable.docx"
inputFile = "./Data/Template_Docx_5.docx"

# Create Word document.
document = Document()

# Load the file from disk.
document.LoadFromFile(inputFile)

# Get the table from Word document.
table = document.Sections[0].Tables[0] if isinstance(
    document.Sections[0].Tables[0], Table) else None

# Change the paragraph setting to keep them together.
for i in range(table.Rows.Count):
    row = table.Rows.get_Item(i)
    for j in range(row.Cells.Count):
        cell = row.Cells.get_Item(j)
        for k in range(cell.Paragraphs.Count):
            p = cell.Paragraphs.get_Item(k)
            p.Format.KeepFollow = True

# Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
