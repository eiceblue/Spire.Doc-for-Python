from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Bookmark.docx"
outputFile = "ReplaceWithTable.docx"
#Load the document
doc = Document()
doc.LoadFromFile(inputFile)

#Create a table
table = Table(doc, True)

#Create data
rowsCount = 4
colsCount = 5
data = [[0 for _ in range(colsCount)] for _ in range(rowsCount)]
data[0] = ["Name", "Capital", "Continent", "Area", "Population"]
data[1] = ["Argentina", "Buenos Aires", "South America", "2777815", "32300003"]
data[2] = ["Bolivia", "La Paz", "South America", "1098575", "7300000"]
data[3] = ["Brazil", "Brasilia", "South America", "8511196", "150400000"]

    
table.ResetCells(rowsCount, colsCount)

for i in range(rowsCount):
    for j in range(colsCount):
        table.Rows[i].Cells[j].AddParagraph().AppendText(data[i][j])

#Get the specific bookmark by its name
navigator = BookmarksNavigator(doc)
navigator.MoveToBookmark("Test")

#Create a TextBodyPart instance and add the table to it
part = TextBodyPart(doc)
part.BodyItems.Add(table)

#Replace the current bookmark content with the TextBodyPart object
navigator.ReplaceBookmarkContent(part)

#Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()