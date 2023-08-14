from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/GetDiagonalBorderOfCell.docx"
outputFile = "GetDiagonalBorder.txt"

# Load Word from disk
doc = Document()
doc.LoadFromFile(inputFile)

# Get the first section
section = doc.Sections[0]

# Get the first table in the section
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

stringBuilder = ''

# Get the setting of the diagonal border of table cell
bs_UP = table.Rows[0].Cells[0].CellFormat.Borders.DiagonalUp.BorderType
stringBuilder += "DiagonalUp border type of table cell(0,0) is " + \
    bs_UP.name
stringBuilder += "\n"
color_UP = table.Rows[0].Cells[0].CellFormat.Borders.DiagonalUp.Color
stringBuilder += "DiagonalUp border color of table cell(0,0) is " + \
    color_UP.ToString()
stringBuilder += "\n"
width_UP = table.Rows[0].Cells[0].CellFormat.Borders.DiagonalUp.LineWidth
stringBuilder += "Line width of DiagonalUp border of table cell(0,0) is " + str(
    width_UP)
stringBuilder += "\n"
bs_Down = table.Rows[0].Cells[0].CellFormat.Borders.DiagonalDown.BorderType
stringBuilder += "DiagonalDown border type of table cell(0,0) is " + \
    bs_Down.name
stringBuilder += "\n"
color_Down = table.Rows[0].Cells[0].CellFormat.Borders.DiagonalDown.Color
stringBuilder += "DiagonalDown border color of table cell(0,0) is " + \
    color_Down.ToString()
stringBuilder += "\n"
width_Down = table.Rows[0].Cells[0].CellFormat.Borders.DiagonalDown.LineWidth
stringBuilder += "DiagonalDown border color of table cell(0,0) is " + str(
    width_Down)
stringBuilder += "\n"

# Save to txt file
f2=open(outputFile,'w', encoding='UTF-8')
f2.write(stringBuilder)
f2.close()
doc.Close()
