from spire.doc import *
from spire.doc.common import *

outputFile = "InsertTableIntoTextBox.docx"

#Create a new document
doc = Document()

#Add a section
section = doc.AddSection()

#Add a paragraph to the section
paragraph = section.AddParagraph()

#Add a textbox to the paragraph
textbox = paragraph.AppendTextBox(300, 100)

#Set the position of the textbox
textbox.Format.HorizontalOrigin = HorizontalOrigin.Page
textbox.Format.HorizontalPosition = 140
textbox.Format.VerticalOrigin = VerticalOrigin.Page
textbox.Format.VerticalPosition = 50

#Add text to the textbox
textboxParagraph = textbox.Body.AddParagraph()
textboxRange = textboxParagraph.AppendText("Table 1")
textboxRange.CharacterFormat.FontName = "Arial"

#Insert table to the textbox
table = textbox.Body.AddTable(True)

#Specify the number of rows and columns of the table
table.ResetCells(4, 4)
data = [["Name", "Age", "Gender", "ID"], ["John", "28", "Male", "0023"], ["Steve", "30", "Male", "0024"], ["Lucy", "26", "female", "0025"]]

#Add data to the table 
for i in range(0, 4):
    for j in range(0, 4):
        tableRange = table.Rows[i].Cells[j].AddParagraph().AppendText(data[i][j])
        tableRange.CharacterFormat.FontName = "Arial"

#Apply style to the table
table.ApplyStyle(DefaultTableStyle.TableColorful2)

#Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
