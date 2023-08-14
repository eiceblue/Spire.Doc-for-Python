from spire.doc import *
from spire.doc.common import *


def _MoidyTableFormat(table):
    # Set table width
    table.PreferredWidth = PreferredWidth(WidthType.Twip, int(6000))

    # Apply style for table
    table.ApplyStyle(DefaultTableStyle.ColorfulGridAccent3)

    # Set table padding
    table.TableFormat.Paddings.SetAll(5)

    # Set table title and description
    table.Title = "Spire.Doc for Python"
    table.TableDescription = "Spire.Doc for Python is a professional Word Python library"


def _ModifyRowFormat(table):
    # Set cell spacing
    table.Rows[0].RowFormat.CellSpacing = 2

    # Set row height
    table.Rows[1].HeightType = TableRowHeightType.Exactly
    table.Rows[1].Height = 20

    # Set background color
    table.Rows[2].RowFormat.BackColor = Color.get_DarkSeaGreen()


def _ModifyCellFormat(table):
    # Set alignment
    table.Rows[0].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle
    table.Rows[0].Cells[0].Paragraphs[0].Format.HorizontalAlignment = HorizontalAlignment.Center

    # Set background color
    table.Rows[1].Cells[0].CellFormat.BackColor = Color.get_DarkSeaGreen()

    # Set cell border
    table.Rows[2].Cells[0].CellFormat.Borders.BorderType(
        BorderStyle.Single)
    table.Rows[2].Cells[0].CellFormat.Borders.LineWidth(1)
    table.Rows[2].Cells[0].CellFormat.Borders.Left.Color = Color.get_Red()
    table.Rows[2].Cells[0].CellFormat.Borders.Right.Color = Color.get_Red()
    table.Rows[2].Cells[0].CellFormat.Borders.Top.Color = Color.get_Red()
    table.Rows[2].Cells[0].CellFormat.Borders.Bottom.Color = Color.get_Red()

    # Set text direction
    table.Rows[3].Cells[0].CellFormat.TextDirection = TextDirection.RightToLeft


outputFile = "ModifyTableFormat.docx"
inputFile = "./Data/ModifyTableFormat.docx"

# Load Word document from disk
document = Document()
document.LoadFromFile(inputFile)

# Get the first section
section = document.Sections[0]

# Get tables
tb1 = section.Tables[0] if isinstance(
    section.Tables[0], Table) else None
tb2 = section.Tables[1] if isinstance(
    section.Tables[1], Table) else None
tb3 = section.Tables[2] if isinstance(
    section.Tables[2], Table) else None

_MoidyTableFormat(tb1)
_ModifyRowFormat(tb2)
_ModifyCellFormat(tb3)

document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
