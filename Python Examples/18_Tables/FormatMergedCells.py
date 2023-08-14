from spire.doc import *
from spire.doc.common import *

def AddTable(section):
    table = section.AddTable(True)
    table.ResetCells(4, 3)
    #Table data
    dt = [
        ["Product", "", "Price"],
        ["Spire.Doc", "Pro Edition", "$799"],
        ["", "Standard Edition", "$599"],
        ["", "Free Edition", "$0"]
    ]

    for r in range(len(dt)):
        dataRow = table.Rows[r]
        dataRow.Height = 20
        dataRow.HeightType = TableRowHeightType.Exactly
        dataRow.RowFormat.BackColor = Color.Empty()
        for c in range(len(dt[r])):
            if not str(dt[r][c]) == '':
                textRange = dataRow.Cells[c].AddParagraph().AppendText(dt[r][c])
                textRange.CharacterFormat.FontName = "Arial"


    return table
  

outputFile = "FormatMergedCells.docx"

# Create word document
document = Document()

# Add a new section
section = document.AddSection()

# Add a new table
table = AddTable(section)

# Create a new style
style = ParagraphStyle(document)
style.Name = "Style"
style.CharacterFormat.TextColor = Color.get_DeepSkyBlue()
style.CharacterFormat.Italic = True
style.CharacterFormat.Bold = True
style.CharacterFormat.FontSize = 13
document.Styles.Add(style)

# Merge cell horizontally
table.ApplyHorizontalMerge(0, 0, 1)
# Apply style
table.Rows[0].Cells[0].Paragraphs[0].ApplyStyle(style.Name)
# Set vertical and horizontal alignment
table.Rows[0].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle
table.Rows[0].Cells[0].Paragraphs[0].Format.HorizontalAlignment = HorizontalAlignment.Center

# Merge cell vertically
table.ApplyVerticalMerge(0, 1, 3)
# Apply style
table.Rows[1].Cells[0].Paragraphs[0].ApplyStyle(style.Name)
# Set vertical and horizontal alignment
table.Rows[1].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle
table.Rows[1].Cells[0].Paragraphs[0].Format.HorizontalAlignment = HorizontalAlignment.Left
# Set column width
table.Rows[1].Cells[0].SetCellWidth(20, CellWidthType.Percentage)
# Save and launch document
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()


