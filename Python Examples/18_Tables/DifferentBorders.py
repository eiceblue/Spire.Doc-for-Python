from spire.doc import *
from spire.doc.common import *

    
def setTableBorders(table):
    table.Format.Borders.BorderType=BorderStyle.Single
    table.Format.Borders.LineWidth=3.0
    table.Format.Borders.Color=Color.get_Red()
def setCellBorders(tableCell):
    tableCell.CellFormat.Borders.BorderType=BorderStyle.DotDash
    tableCell.CellFormat.Borders.LineWidth=1.0
    tableCell.CellFormat.Borders.Color=Color.get_Green()

inputFile = "./Data/TableSample.docx"
outputFile = "DifferentBorders.docx"

# Open a Word document as template
document = Document()
document.LoadFromFile(inputFile)

table = document.Sections[0].Tables[0] if isinstance(document.Sections[0].Tables[0], Table) else None

# Set borders of table
setTableBorders(table)

# Set borders of cell
setCellBorders(table.Rows[2].Cells[0])

# Save and launch document
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()




