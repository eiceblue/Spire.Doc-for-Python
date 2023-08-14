import math
from spire.doc import *
from spire.doc.common import *

def addTable(section):
    header = ["Name", "Capital", "Continent", "Area", "Population"]
    data = [["Argentina", "Buenos Aires", "South America", "2777815", "32300003"], ["Bolivia", "La Paz", "South America", "1098575", "7300000"], ["Brazil", "Brasilia", "South America", "8511196", "150400000"], ["Canada", "Ottawa", "North America", "9976147", "26500000"], ["Chile", "Santiago", "South America", "756943", "13200000"], ["Colombia", "Bagota", "South America", "1138907", "33000000"], ["Cuba", "Havana", "North America", "114524", "10600000"], ["Ecuador", "Quito", "South America", "455502", "10600000"], ["El Salvador", "San Salvador", "North America", "20865", "5300000"], ["Guyana", "Georgetown", "South America", "214969", "800000"], ["Jamaica", "Kingston", "North America", "11424", "2500000"], ["Mexico", "Mexico City", "North America", "1967180", "88600000"], ["Nicaragua", "Managua", "North America", "139000", "3900000"], ["Paraguay", "Asuncion", "South America", "406576", "4660000"], ["Peru", "Lima", "South America", "1285215", "21600000"], ["United States of America", "Washington", "North America", "9363130", "249200000"], ["Uruguay", "Montevideo", "South America", "176140", "3002000"], ["Venezuela", "Caracas", "South America", "912047", "19700000"]]
    table = section.AddTable(True)
    table.ResetCells(len(data) + 1, len(header))

    # ***************** First Row *************************
    row = table.Rows[0]
    row.IsHeader = True
    row.Height = 20 #unit: point, 1point = 0.3528 mm
    row.HeightType = TableRowHeightType.Exactly
    row.RowFormat.BackColor = Color.get_Gray()
    i = 0
    while i < len(header):
        row.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle
        p = row.Cells[i].AddParagraph()
        p.Format.HorizontalAlignment = HorizontalAlignment.Center
        txtRange = p.AppendText(header[i])
        txtRange.CharacterFormat.Bold = True
        i += 1

    r = 0
    while r < len(data):
        dataRow = table.Rows[r + 1]
        dataRow.Height = 20
        dataRow.HeightType = TableRowHeightType.Exactly
        dataRow.RowFormat.BackColor = Color.Empty()
        c = 0
        while c < len(data[r]):
            dataRow.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle
            dataRow.Cells[c].AddParagraph().AppendText(data[r][c])
            c += 1
        r += 1

    for j in range(1, table.Rows.Count):
        if math.fmod(j, 2) == 0:
            row2 = table.Rows[j]
            for f in range(row2.Cells.Count):
                row2.Cells[f].CellFormat.BackColor = Color.get_LightBlue()

     
outputFile = "CreateTable.docx"

# Open a blank Word document as template
document = Document()

section = document.AddSection()
addTable(section)

# Save docx file
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()


