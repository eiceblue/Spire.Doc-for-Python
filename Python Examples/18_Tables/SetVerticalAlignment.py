from spire.doc import *
from spire.doc.common import *

outputFile = "SetVerticalAlignment.docx"
inputFile = "./Data/E-iceblue.png"

# Create a new Word document and add a new section
doc = Document()
section = doc.AddSection()

# Add a table with 3 columns and 3 rows
table = section.AddTable(True)
table.ResetCells(3, 3)

# Merge rows
table.ApplyVerticalMerge(0, 0, 2)

# Set the vertical alignment for each cell, default is top
table.Rows[0].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle
table.Rows[0].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Top
table.Rows[0].Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Top
table.Rows[1].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle
table.Rows[1].Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Middle
table.Rows[2].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Bottom
table.Rows[2].Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Bottom

# Inset a picture into the table cell
paraPic = table.Rows[0].Cells[0].AddParagraph()
pic = paraPic.AppendPicture(inputFile)

# Create data
data = [["", "Spire.Office", "Spire.DataExport"], [
    "", "Spire.Doc", "Spire.DocViewer"], ["", "Spire.XLS", "Spire.PDF"]]

# Append data to table
for r in range(0, 3):
    dataRow = table.Rows[r]
    dataRow.Height = 50
    for c in range(0, 3):
        if c == 1:
            par = dataRow.Cells[c].AddParagraph()
            par.AppendText(data[r][c])
            dataRow.Cells[c].SetCellWidth((section.PageSetup.ClientWidth) / 2, CellWidthType.Point)
        if c == 2:
            par = dataRow.Cells[c].AddParagraph()
            par.AppendText(data[r][c])
            dataRow.Cells[c].SetCellWidth((section.PageSetup.ClientWidth) / 2, CellWidthType.Point)

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
