from spire.doc import *
from spire.doc.common import *

outputFile = "SetTableStyleAndBorder.docx"
inputFile = "./Data/TableSample.docx"

# Create a document and load file
document = Document()
document.LoadFromFile(inputFile)

section = document.Sections[0]

# Get the first table
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

# Apply the table style
table.ApplyStyle(DefaultTableStyle.ColorfulList)

# Set right border of table
table.Format.Borders.Right.BorderType = BorderStyle.Hairline
table.Format.Borders.Right.LineWidth = 1.0
table.Format.Borders.Right.Color = Color.get_Red()

# Set top border of table
table.Format.Borders.Top.BorderType = BorderStyle.Hairline
table.Format.Borders.Top.LineWidth = 1.0
table.Format.Borders.Top.Color = Color.get_Green()

# Set left border of table
table.Format.Borders.Left.BorderType = BorderStyle.Hairline
table.Format.Borders.Left.LineWidth = 1.0
table.Format.Borders.Left.Color = Color.get_Yellow()

# Set bottom border is none
table.Format.Borders.Bottom.BorderType = BorderStyle.DotDash

# Set vertical and horizontal border
table.Format.Borders.Vertical.BorderType = BorderStyle.Dot
table.Format.Borders.Horizontal.BorderType = BorderStyle.none
table.Format.Borders.Vertical.Color = Color.get_Orange()

# Save the file and launch it
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
