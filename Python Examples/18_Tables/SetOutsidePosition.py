from spire.doc import *
from spire.doc.common import *


outputFile = "SetOutsidePosition.docx"
inputFile = "./Data/Spire.Doc.png"

# Create a new word document and add new section
doc = Document()
sec = doc.AddSection()

# Get header
header = doc.Sections[0].HeadersFooters.Header

# Add new paragraph on header and set HorizontalAlignment of the paragraph as left
paragraph = header.AddParagraph()
paragraph.Format.HorizontalAlignment = HorizontalAlignment.Left

# Load an image for the paragraph
headerimage = paragraph.AppendPicture(inputFile)

# Add a table of 4 rows and 2 columns
table = header.AddTable()
table.ResetCells(4, 2)

# Set the position of the table to the right of the image
table.Format.WrapTextAround = True
table.Format.Positioning.HorizPositionAbs = HorizontalPosition.Outside
table.Format.Positioning.VertRelationTo = VerticalRelation.Margin
table.Format.Positioning.VertPosition = 43

# Add contents for the table
data = [["Spire.Doc.left", "Spire XLS.right"], ["Spire.Presentatio.left", "Spire.PDF.right"], [
    "Spire.DataExport.left", "Spire.PDFViewe.right"], ["Spire.DocViewer.left", "Spire.BarCode.right"]]

for r in range(0, 4):
    dataRow = table.Rows[r]
    for c in range(0, 2):
        if c == 0:
            par = dataRow.Cells[c].AddParagraph()
            par.AppendText(data[r][c])
            par.Format.HorizontalAlignment = HorizontalAlignment.Left
            dataRow.Cells[c].SetCellWidth(180,CellWidthType.Point)
        else:
            par = dataRow.Cells[c].AddParagraph()
            par.AppendText(data[r][c])
            par.Format.HorizontalAlignment = HorizontalAlignment.Right
            dataRow.Cells[c].SetCellWidth(180,CellWidthType.Point)

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
