from spire.doc import *
from spire.doc.common import *

outputFile =  "ExtractTextFromTextBoxes.txt"
inputFile = "./Data/ExtractTextFromTextBoxes.docx"

#Create a Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)

#Verify whether the document contains a textbox or not.
if document.TextBoxes.Count > 0:
    with open(outputFile,'w') as sw:
        #Traverse the document.
        for i in range(document.Sections.Count):
            section = document.Sections.get_Item(i)
            for j in range(section.Paragraphs.Count):
                p = section.Paragraphs.get_Item(j)
                for k in range(p.ChildObjects.Count):
                    obj = p.ChildObjects.get_Item(k)
                    if obj.DocumentObjectType == DocumentObjectType.TextBox:
                        textbox = obj if isinstance(obj, TextBox) else None
                        for x in range(textbox.ChildObjects.Count):
                            objt = textbox.ChildObjects.get_Item(x)
                            #Extract text from paragraph in TextBox.
                            if objt.DocumentObjectType == DocumentObjectType.Paragraph:
                                sw.write(( objt if isinstance(objt, Paragraph) else None).Text)
                            #Extract text from Table in TextBox.
                            if objt.DocumentObjectType == DocumentObjectType.Table:
                                table = objt if isinstance(objt, Table) else None
                                for i in range(table.Rows.Count):
                                    row = table.Rows[i]
                                    for j in range(row.Cells.Count):
                                        cell = row.Cells[j]
                                        for k in range(cell.Paragraphs.Count):
                                            paragraph = cell.Paragraphs.get_Item(k)
                                            sw.write(paragraph.Text)
document.Close()
