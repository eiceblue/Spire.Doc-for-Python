from spire.doc import *
from spire.doc.common import *

outputFile = "CreateTableDirectly.docx"

# Create a Word document
doc = Document()

# Add a section
section = doc.AddSection()

# Create a table
table = Table(doc)
# Set the width of table
table.PreferredWidth = PreferredWidth(WidthType.Percentage, int(100))
# Set the border of table
table.TableFormat.Borders.BorderType = BorderStyle.Single

# Create a table row
row = TableRow(doc)
row.Height = 50.0
table.Rows.Add(row)

# Create a table cell
cell1 = TableCell(doc)
# Add a paragraph
para1 = cell1.AddParagraph()
# Append text in the paragraph
para1.AppendText("Row 1, Cell 1")
# Set the horizontal alignment of paragrah
para1.Format.HorizontalAlignment = HorizontalAlignment.Center
# Set the background color of cell
cell1.CellFormat.BackColor = Color.get_CadetBlue()
# Set the vertical alignment of paragraph
cell1.CellFormat.VerticalAlignment = VerticalAlignment.Middle
row.Cells.Add(cell1)

# Create a table cell
cell2 = TableCell(doc)
para2 = cell2.AddParagraph()
para2.AppendText("Row 1, Cell 2")
para2.Format.HorizontalAlignment = HorizontalAlignment.Center
cell2.CellFormat.BackColor = Color.get_CadetBlue()
cell2.CellFormat.VerticalAlignment = VerticalAlignment.Middle
row.Cells.Add(cell2)

# Add the table in the section
section.Tables.Add(table)

# Save the document
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()
