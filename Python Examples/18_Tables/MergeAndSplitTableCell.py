from spire.doc import *
from spire.doc.common import *

outputFile = "MergeAndSplitTableCell.docx"
inputFile = "./Data/TableSample.docx"

# Create a document and load file from disk
document = Document()
document.LoadFromFile(inputFile)
section = document.Sections[0]
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None
# The method shows how to merge cell horizontally
table.ApplyHorizontalMerge(6, 2, 3)
# The method shows how to merge cell vertically
table.ApplyVerticalMerge(2, 4, 5)
# The method shows how to split the cell
table.Rows[8].Cells[3].SplitCell(2, 2)
# Save to file and launch it
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
