import math
from spire.doc import *
from spire.doc.common import *

outputFile = "RepeatRowOnEachPage.docx"

# Create word document
document = Document()

# Create a new section
section = document.AddSection()

# Create a table width default borders
table = section.AddTable(True)
# Set table with to 100%
width = PreferredWidth(WidthType.Percentage, 100)
table.PreferredWidth = width

# Add a new row
row = table.AddRow()
# Set the row as a table header
row.IsHeader = True
# Set the backcolor of row
row.RowFormat.BackColor = Color.get_LightGray()
# Add a new cell for row
cell = row.AddCell()
cell.SetCellWidth(100, CellWidthType.Percentage)
# Add a paragraph for cell to put some data
parapraph = cell.AddParagraph()
# Add text
parapraph.AppendText("Row Header 1")
# Set paragraph horizontal center alignment
parapraph.Format.HorizontalAlignment = HorizontalAlignment.Center

row = table.AddRow(False, 1)
row.IsHeader = True
row.RowFormat.BackColor = Color.get_Ivory()
# Set row height
row.Height = 30
cell = row.Cells[0]
cell.SetCellWidth(100, CellWidthType.Percentage)
# Set cell vertical middle alignment
cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
# Add a paragraph for cell to put some data
parapraph = cell.AddParagraph()
# Add text
parapraph.AppendText("Row Header 2")
parapraph.Format.HorizontalAlignment = HorizontalAlignment.Center

# Add many common rows
for i in range(0, 70):
    row = table.AddRow(False, 2)
    cell = row.Cells[0]
    # Set cell width to 50% of table width
    cell.SetCellWidth(50, CellWidthType.Percentage)
    cell.AddParagraph().AppendText("Column 1 Text")
    cell = row.Cells[1]
    cell.SetCellWidth(50, CellWidthType.Percentage)
    cell.AddParagraph().AppendText("Column 2 Text")
# Set cell backcolor
for j in range(1, table.Rows.Count):
    if math.fmod(j, 2) == 0:
        row2 = table.Rows[j]
        for f in range(row2.Cells.Count):
            row2.Cells[f].CellFormat.BackColor = Color.get_LightBlue()

# Save file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
