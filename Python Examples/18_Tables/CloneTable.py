from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/TableTemplate.docx"
outputFile = "CloneTable.docx"

# Load the document
doc = Document()
doc.LoadFromFile(inputFile)

# Get the first section
se = doc.Sections[0]

# Get the first table
original_Table = se.Tables[0]

# Copy the existing table to copied_Table via Table.clone()
copied_Table = original_Table.Clone()
st = ["Spire.Presentation for Python", "A professional " + "PowerPoint® compatible library that enables developers to create, read, " +
       "write, modify, convert and Print PowerPoint documents on python platforms."]

# Get the last row of table
lastRow = copied_Table.Rows[copied_Table.Rows.Count - 1]

# Change last row data
i = 0
while i < lastRow.Cells.Count - 1:
    lastRow.Cells[i].Paragraphs[0].Text = st[i]
    i += 1
    
# Add copied_Table in section
se.Tables.Add(copied_Table)

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
