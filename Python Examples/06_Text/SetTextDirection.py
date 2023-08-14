from spire.doc import *
from spire.doc.common import *

outputFile = "SetTextDirection.docx"

#Create a new document
doc = Document()
#Add the first section
section1 = doc.AddSection()
#Set text direction for all text in a section
section1.TextDirection = TextDirection.RightToLeft
#Set Font Style and Size
style = ParagraphStyle(doc)
style.Name = "FontStyle"
style.CharacterFormat.FontName = "Arial"
style.CharacterFormat.FontSize = 15
doc.Styles.Add(style)
#Add two paragraphs and apply the font style
p = section1.AddParagraph()
p.AppendText("Only Spire.Doc, no Microsoft Office automation")
p.ApplyStyle(style.Name)
p = section1.AddParagraph()
p.AppendText("Convert file documents with high quality")
p.ApplyStyle(style.Name)
#Set text direction for a part of text
#Add the second section
section2 = doc.AddSection()
#Add a table
table = section2.AddTable()
table.ResetCells(1, 1)
cell = table.Rows[0].Cells[0]
table.Rows[0].Height = 150
table.Rows[0].Cells[0].SetCellWidth(10, CellWidthType.Point)
#Set vertical text direction of table
cell.CellFormat.TextDirection = TextDirection.RightToLeftRotated
cell.AddParagraph().AppendText("This is vertical style")
#Add a paragraph and set horizontal text direction
p = section2.AddParagraph()
p.AppendText("This is horizontal style")
p.ApplyStyle(style.Name)
#Save the document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()