from spire.doc import *
from spire.doc.common import *

outputFile = "CreateVerticalTable.docx"

# Create Word document.
document = Document()

# Add a new section.
section = document.AddSection()

# Add a table with rows and columns and set the text for the table.
table = section.AddTable()
table.ResetCells(1, 1)
cell = table.Rows[0].Cells[0]
table.Rows[0].Height = 150
cell.AddParagraph().AppendText("Draft copy in vertical style")

# Set the TextDirection for the table to RightToLeftRotated.
cell.CellFormat.TextDirection = TextDirection.RightToLeftRotated

# Set the table format.
table.Format.WrapTextAround = True
table.Format.Positioning.VertRelationTo = VerticalRelation.Page
table.Format.Positioning.HorizRelationTo = HorizontalRelation.Page
table.Format.Positioning.HorizPosition = section.PageSetup.PageSize.Width - table.Width
table.Format.Positioning.VertPosition = 200

# Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
