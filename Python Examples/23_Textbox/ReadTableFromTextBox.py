from spire.doc import *
from spire.doc.common import *

outputFile = "ReadTableFromTextBox.txt"
inputFile = "./Data/TextBoxTable.docx"

#Load the document
doc = Document()
doc.LoadFromFile(inputFile)
#Get the first textbox
textbox = doc.TextBoxes[0]
#Get the first table in the textbox
table = textbox.Body.Tables[0] if isinstance(textbox.Body.Tables[0], Table) else None
tempStr = ''
#Loop through the paragraphs of the table cells and extract them to a .txt file
for i in range(table.Rows.Count):
    row = table.Rows.get_Item(i)
    for j in range(row.Cells.Count):
        cell = row.Cells.get_Item(j)
        for k in range(cell.Paragraphs.Count):
            paragraph = cell.Paragraphs.get_Item(k)
            tempStr += paragraph.Text + "\t"
    tempStr += "\r\n"
#Save to TXT file and launch it
with open(outputFile,'w') as fp:
            fp.write(tempStr)
doc.Close()