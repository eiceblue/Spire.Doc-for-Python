from spire.doc import *
from spire.doc.common import *

outputFile = "ReplaceTextWithMergeField.docx"
inputFile = "./Data/SampleB_2.docx"

#Open a Word document
document = Document()
document.LoadFromFile(inputFile)

#Find the text that will be replaced
ts = document.FindString("Test", True, True)

tr = ts.GetAsOneRange()

#Get the paragraph
par = tr.OwnerParagraph

#Get the index of the text in the paragraph
index = par.ChildObjects.IndexOf(tr)

#Create a new field
field = MergeField(document)
field.FieldName = "MergeField"

#Insert field at specific position
par.ChildObjects.Insert(index, field)

#Remove the text
par.ChildObjects.Remove(tr)

#Save to file
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()

