from spire.doc import *

def add_cell_content(cell, content):
    # Add a new paragraph to the table cell.
    paragraph = cell.AddParagraph()

    # Append the specified text content to the paragraph inside the cell.
    paragraph.AppendText(content)

    # Center-align the text horizontally within the paragraph.
    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center

    # Set the background (shading) color of the table cell to CadetBlue.
    cell.CellFormat.Shading.BackgroundPatternColor = Color.get_CadetBlue()

    # Vertically center the content within the table cell.
    cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle

# Create a new empty document instance.
doc = Document()

# Create a document navigator to help navigate and modify the document content.
navigator = DocumentNavigator(doc)

# Start creating a new table and get its reference.
table = navigator.StartTable()

# Initialize the table with 2 rows and 2 columns.
table.ResetCells(2, 2)

# Set the table width to 100% of the available page width.
table.PreferredWidth = PreferredWidth(WidthType.Percentage, 100)

# Apply a single black border with line width 1 to all sides of the table.
table.Format.Borders.SetBordersAttribute(BorderStyle.Single, 1, Color.get_Black())

# Set the height of the first row to 30 points.
table.FirstRow.Height = 30.0

# Get a reference to the first cell in the first row (row 0, column 0).
cell1 = table.Rows[0].Cells[0]

# Add content to the first cell of the first row.
add_cell_content(cell1, "Row 1, Cell 1")

# Get a reference to the second cell in the first row (row 0, column 1).
cell2 = table.Rows[0].Cells[1]

# Add content to the second cell of the first row.
add_cell_content(cell2, "Row 1, Cell 2")

# Get a reference to the first cell in the second row (row 1, column 0).
cell3 = table.Rows[1].Cells[0]

# Add content to the first cell of the second row.
add_cell_content(cell3, "Row 2, Cell 1")

# Get a reference to the second cell in the second row (row 1, column 1).
cell4 = table.Rows[1].Cells[1]

# Add content to the second cell of the second row.
add_cell_content(cell4, "Row 2, Cell 2")

# Finalize the table creation.
navigator.EndTable()

# Delete the first row (row index 0) of the first table (table index 0) in the document.
navigator.DeleteRow(0, 0)

# Save the modified document to a new file named "OperateTableRow.docx" in DOCX format.
doc.SaveToFile("OperateTableRow.docx", FileFormat.Docx)

# Close the document to release internal resources.
doc.Close()

# Explicitly dispose of the document object to free memory immediately.
doc.Dispose()