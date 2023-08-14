from io import FileIO
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ReplaceTextInTable.docx"
outputFile = "GetRowCellIndex.txt"

# Load Word from disk
doc = Document()
doc.LoadFromFile(inputFile)

# Get the first section
section = doc.Sections[0]

# Get the first table in the section
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

content = ''

# Get table collections
collections = section.Tables

# Get the table index
tableIndex = collections.IndexOf(table)

# Get the index of the last table row
row = table.LastRow
rowIndex = row.GetRowIndex()

# Get the index of the last table cell
cell = row.LastChild if isinstance(row.LastChild, TableCell) else None
cellIndex = cell.GetCellIndex()

# Append these information into content
content += "Table index is " + str(tableIndex)
content += "\n"
content += "Row index is " + str(rowIndex)
content += "\n"
content += "Cell index is " + str(cellIndex)
content += "\n"

# Save to txt file
f2=open(outputFile,'w', encoding='UTF-8')
f2.write(content)
f2.close()
doc.Close()
