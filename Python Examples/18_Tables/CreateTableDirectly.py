from spire.doc import *
from spire.doc.common import *

outputFile = "CreateTableDirectly.docx"

# Create a Word document
doc = Document()

# Add a section
section = doc.AddSection()

#Create a table 
table = Table(doc)
table.ResetCells(1,2)

#Set the width of table
table.PreferredWidth = PreferredWidth(WidthType.Percentage, int(100))

#Set the border of table
table.Format.Borders.BorderType = BorderStyle.Single

#Create a table row
row = table.Rows[0]
row.Height = 50.0

#Create a table cell
cell1 = table.Rows[0].Cells[0]

#Add a paragraph
para1 = cell1.AddParagraph()
#Append text in the paragraph

para1.AppendText("Row 1, Cell 1")
#Set the horizontal alignment of paragrah
para1.Format.HorizontalAlignment = HorizontalAlignment.Center

#Set the background color of cell
cell1.CellFormat.Shading.BackgroundPatternColor = Color.get_CadetBlue()

#Set the vertical alignment of paragraph
cell1.CellFormat.VerticalAlignment = VerticalAlignment.Middle

#Create a table cell
cell2 = table.Rows[0].Cells[1]
para2 = cell2.AddParagraph()
para2.AppendText("Row 1, Cell 2")
para2.Format.HorizontalAlignment = HorizontalAlignment.Center
cell2.CellFormat.Shading.BackgroundPatternColor = Color.get_CadetBlue()
cell2.CellFormat.VerticalAlignment = VerticalAlignment.Middle
row.Cells.Add(cell2)

#Add the table in the section
section.Tables.Add(table)

#Save the document
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()
