
from io import *
from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/CellMergeStatus.docx"
outputFile = "CellMergeStatus.txt"
doc = Document()
doc.LoadFromFile(inputFile)

# Get the first section
section = doc.Sections[0]

# Get the first table in the section
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

content = ''
for i in range(table.Rows.Count):
    tableRow = table.Rows[i]
    for j in range(tableRow.Cells.Count):
        tableCell = tableRow.Cells[j]
        verticalMerge = tableCell.CellFormat.VerticalMerge
        horizontalMerge = tableCell.GridSpan
        if verticalMerge == CellMerge.none and horizontalMerge == 1:
            content += "Row " + str(i) + ", cell " + str(j) + ": "
            content += "This cell isn't merged."
            content += "\n"
        else:
            content += "Row " + str(i) + ", cell " + str(j) + ": "
            content += "This cell is merged."
            content += "\n"
    content += "\n"

# Save and launch document
f2=open(outputFile,'w', encoding='UTF-8')
f2.write(content)
f2.close()
doc.Close()
