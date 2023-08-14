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
table.TableFormat.Borders.Right.BorderType = BorderStyle.Hairline
table.TableFormat.Borders.Right.LineWidth = 1.0
table.TableFormat.Borders.Right.Color = Color.get_Red()

# Set top border of table
table.TableFormat.Borders.Top.BorderType = BorderStyle.Hairline
table.TableFormat.Borders.Top.LineWidth = 1.0
table.TableFormat.Borders.Top.Color = Color.get_Green()

# Set left border of table
table.TableFormat.Borders.Left.BorderType = BorderStyle.Hairline
table.TableFormat.Borders.Left.LineWidth = 1.0
table.TableFormat.Borders.Left.Color = Color.get_Yellow()

# Set bottom border is none
table.TableFormat.Borders.Bottom.BorderType = BorderStyle.DotDash

# Set vertical and horizontal border
table.TableFormat.Borders.Vertical.BorderType = BorderStyle.Dot
table.TableFormat.Borders.Horizontal.BorderType = BorderStyle.none
table.TableFormat.Borders.Vertical.Color = Color.get_Orange()

# Save the file and launch it
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
