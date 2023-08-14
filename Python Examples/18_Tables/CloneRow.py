from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/TableTemplate.docx"
outputFile = "CloneRow.docx"

doc = Document()
doc.LoadFromFile(inputFile)

# Get the first section
se = doc.Sections[0]

# Get the first row of the first table
firstRow = se.Tables[0].Rows[0]

# Copy the first row to clone_FirstRow via TableRow.clone()
clone_FirstRow = firstRow.Clone()

se.Tables[0].Rows.Add(clone_FirstRow)
# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
