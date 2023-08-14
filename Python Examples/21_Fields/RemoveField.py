from spire.doc import *
from spire.doc.common import *

outputFile = "RemoveField.docx"
inputFile = "./Data/IfFieldSample.docx"

#Open a Word document
document = Document()
document.LoadFromFile(inputFile)
#Get the first field
field = document.Fields[0]
#Get the paragraph of the field
par = field.OwnerParagraph
#Get the index of the  field
index = par.ChildObjects.IndexOf(field)
#Remove if field via index
par.ChildObjects.RemoveAt(index)
#Save doc file
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()

