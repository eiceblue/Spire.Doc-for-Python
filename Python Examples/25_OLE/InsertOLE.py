from spire.doc import *
from spire.doc.common import *

outputFile = "InsertOLE.docx"

#create a document
doc = Document()

#add a section
sec = doc.AddSection()

#add a paragraph
par = sec.AddParagraph()

#load the image
picture = DocPicture(doc)
picture.LoadImage("./Data/Excel.png")

#insert the OLE
obj = par.AppendOleObject("./Data/example.xlsx", picture, OleObjectType.ExcelWorksheet)
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()