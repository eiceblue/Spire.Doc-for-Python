from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/TableSample.docx"
outputFile = "AddOrDeleteRow.docx"

# Create a document
document = Document()
# Load file
document.LoadFromFile(inputFile)
section = document.Sections[0]
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

# Delete the seventh row
table.Rows.RemoveAt(7)

# Add a row and insert it into specific position
row = TableRow(document)
for i in range(table.Rows[0].Cells.Count):
    tc = row.AddCell()
    paragraph = tc.AddParagraph()
    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
    paragraph.AppendText("Added")
table.Rows.Insert(2, row)
# Add a row at the end of table
table.AddRow()

# Save to file and launch it
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
