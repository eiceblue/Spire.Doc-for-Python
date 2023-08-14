from spire.doc import *
from spire.doc.common import *

outputFile = "CreateNestedTable.docx"

# Create a new document
doc = Document()
section = doc.AddSection()

# Add a table
table = section.AddTable(True)
table.ResetCells(2, 2)

# Set column width
table.Rows[0].Cells[0].SetCellWidth(70, CellWidthType.Point)
table.Rows[1].Cells[0].SetCellWidth(70, CellWidthType.Point)
table.AutoFit(AutoFitBehaviorType.AutoFitToWindow)

# Insert content to cells
table.Rows[0].Cells[0].AddParagraph().AppendText("Spire.Doc for Python")
text = "Spire.Doc for Python is a professional Word" + "Python library specifically designed for developers to create," + \
    "read, write, convert and print Word document files" + \
    " with fast and high quality performance."
table.Rows[0].Cells[1].AddParagraph().AppendText(text)

# Add a nested table to cell(first row, second column)
nestedTable = table.Rows[0].Cells[1].AddTable(True)
nestedTable.ResetCells(4, 3)
nestedTable.AutoFit(AutoFitBehaviorType.AutoFitToContents)

# Add content to nested cells
nestedTable.Rows[0].Cells[0].AddParagraph().AppendText("NO.")
nestedTable.Rows[0].Cells[1].AddParagraph().AppendText("Item")
nestedTable.Rows[0].Cells[2].AddParagraph().AppendText("Price")

nestedTable.Rows[1].Cells[0].AddParagraph().AppendText("1")
nestedTable.Rows[1].Cells[1].AddParagraph().AppendText("Pro Edition")
nestedTable.Rows[1].Cells[2].AddParagraph().AppendText("$799")

nestedTable.Rows[2].Cells[0].AddParagraph().AppendText("2")
nestedTable.Rows[2].Cells[1].AddParagraph().AppendText("Standard Edition")
nestedTable.Rows[2].Cells[2].AddParagraph().AppendText("$599")

nestedTable.Rows[3].Cells[0].AddParagraph().AppendText("3")
nestedTable.Rows[3].Cells[1].AddParagraph().AppendText("Free Edition")
nestedTable.Rows[3].Cells[2].AddParagraph().AppendText("$0")

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
