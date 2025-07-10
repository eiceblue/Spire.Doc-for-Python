
from spire.doc import *
from spire.doc.common import *


outputFile = "PageSetup.docx"
#Create Word document.
document = Document()
section = document.AddSection()
#The unit of all measures below is point, 1point = 0.3528 mm.
section.PageSetup.PageSize = PageSize.A4()
section.PageSetup.Margins.Top = 72
section.PageSetup.Margins.Bottom = 72
section.PageSetup.Margins.Left = 89.85
section.PageSetup.Margins.Right = 89.85
#Insert header and footer.
header = section.HeadersFooters.Header
footer = section.HeadersFooters.Footer
#Insert picture and text to header.
headerParagraph = header.AddParagraph()
headerPicture = headerParagraph.AppendPicture("./Data/Header.png")
#Header text.
text = headerParagraph.AppendText("Demo of Spire.Doc")
text.CharacterFormat.FontName = "Arial"
text.CharacterFormat.FontSize = 10
text.CharacterFormat.Italic = True
headerParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right
#Border.
headerParagraph.Format.Borders.Bottom.BorderType = BorderStyle.Single
headerParagraph.Format.Borders.Bottom.Space = 0.05
#Header picture layout - text wrapping.
headerPicture.TextWrappingStyle = TextWrappingStyle.Behind
#Header picture layout - position.
headerPicture.HorizontalOrigin = HorizontalOrigin.Page
headerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left
headerPicture.VerticalOrigin = VerticalOrigin.Page
headerPicture.VerticalAlignment = ShapeVerticalAlignment.Top
#Insert picture to footer.
footerParagraph = footer.AddParagraph()
footerPicture = footerParagraph.AppendPicture("./Data/Footer.png")
#Footer picture layout.
footerPicture.TextWrappingStyle = TextWrappingStyle.Behind
footerPicture.HorizontalOrigin = HorizontalOrigin.Page
footerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left
footerPicture.VerticalOrigin = VerticalOrigin.Page
footerPicture.VerticalAlignment = ShapeVerticalAlignment.Bottom
#Insert page number.
footerParagraph.AppendField("page number", FieldType.FieldPage)
footerParagraph.AppendText(" of ")
footerParagraph.AppendField("number of pages", FieldType.FieldNumPages)
footerParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right
#Border.
footerParagraph.Format.Borders.Top.BorderType = BorderStyle.Single
footerParagraph.Format.Borders.Top.Space = 0.05

header = ["Name", "Capital", "Continent", "Area", "Population"]
data = [["Argentina", "Buenos Aires", "South America", "2777815", "32300003"], ["Bolivia", "La Paz", "South", "1098575", "7300000"], ["Brazil", "Brasilia", "South", "8511196", "150400000"], ["Canada", "Ottawa", "North", "9976147", "26500000"], ["Chile", "Santiago", "South", "756943", "13200000"], ["Colombia", "Bagota", "South", "1138907", "33000000"], ["Cuba", "Havana", "North", "114524", "10600000"], ["Ecuador", "Quito", "South", "455502", "10600000"], ["El Salvador", "San Salvador", "North", "20865", "5300000"], ["Guyana", "Georgetown", "South", "214969", "800000"], ["Jamaica", "Kingston", "North", "11424", "2500000"], ["Mexico", "Mexico City", "North", "1967180", "88600000"], ["Nicaragua", "Managua", "North", "139000", "3900000"], ["Paraguay", "Asuncion", "South", "406576", "4660000"], ["Peru", "Lima", "South", "1285215", "21600000"], ["United States", "Washington", "North", "9363130", "249200000"], ["Uruguay", "Montevideo", "South", "176140", "3002000"], ["Venezuela", "Caracas", "South", "912047", "19700000"]]
table = section.AddTable(True)
table.ResetCells(len(data) + 1, len(header))
# ***************** First Row *************************
row = table.Rows[0]
row.IsHeader = True
row.Height = 20
row.HeightType = TableRowHeightType.Exactly
row.RowFormat.BackColor = Color.get_Gray()

i = 0
while i < row.Cells.Count:
    row.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.get_Gray()
    i += 1
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
    c = 0
    while c < dataRow.Cells.Count:
        dataRow.Cells[c].CellFormat.Shading.BackgroundPatternColor = Color.Empty()
        c += 1
    c = 0
    while c < len(data[r]):
        dataRow.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle
        dataRow.Cells[c].AddParagraph().AppendText(data[r][c])
        c += 1
    r += 1

#Save doc file.
document.SaveToFile(outputFile, FileFormat.Docx2019)
document.Close()
