from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/TableTemplate.docx"
outputFile = "AddPictureToTableCell.docx"

# Load the document
doc = Document()
doc.LoadFromFile(inputFile)

# Get the first table from the first section of the document
table1 = doc.Sections[0].Tables[0]

# Add a picture to the specified table cell and set picture size

picture = table1.Rows[1].Cells[2].Paragraphs[0].AppendPicture("./Data/Spire.Doc.png")

picture.Width = 100
picture.Height = 100

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
