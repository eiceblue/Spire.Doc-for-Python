from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/example.zip"
inputFile_I = "./Data/example.png"
outputFile = "InsertOLEAsIconViaStream.docx"

#Create word document
doc = Document()
#add a section
sec = doc.AddSection()
#add a paragraph
par = sec.AddParagraph()

#ole stream
stream = Stream(inputFile)

#load the image
picture = DocPicture(doc)
picture.LoadImage(inputFile_I)

#insert the OLE from stream
obj = par.AppendOleObject(stream, picture, "zip")

#display as icon
obj.DisplayAsIcon = True
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()
