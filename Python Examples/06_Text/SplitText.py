from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Sample.docx"
outputFile = "SplitText.docx"

#Create a new document and load from file
doc = Document()
doc.LoadFromFile(inputFile)
#Add a column to the first section and set width and spacing
doc.Sections[0].AddColumn(100, 20)
#Add a line between the two columns
doc.Sections[0].PageSetup.ColumnsLineBetween = True
#Save the document
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()