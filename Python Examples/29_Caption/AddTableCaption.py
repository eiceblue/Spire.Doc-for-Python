from spire.doc import *
from spire.doc.common import *

outputFile = "AddTableCaption.docx"
inputFile = "./Data/TableTemplate.docx"

#Create word document
document = Document()

#Load file
document.LoadFromFile(inputFile)

#Get the first table
body = document.Sections[0].Body
table = body.Tables[0] if isinstance(body.Tables[0], Table) else None

#Add caption to the table
table.AddCaption("Table", CaptionNumberingFormat.Number, CaptionPosition.BelowItem)

#Update fields
document.IsUpdateFields = True

#Save the Word document
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
