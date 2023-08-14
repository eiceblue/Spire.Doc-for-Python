from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ReplaceTextWithField.docx"
outputFile = "ReplaceTextWithField.docx"

#Create word document
document = Document()

#Load the document from disk.
document.LoadFromFile(inputFile)

#Find the target text
selection = document.FindString('summary', False, True)

#Get text range
textRange = selection.GetAsOneRange()
#Get it's owner paragraph
ownParagraph:Paragraph = textRange.OwnerParagraph
#Get the index of this text range
rangeIndex = ownParagraph.ChildObjects.IndexOf(textRange)
#Remove the text range
ownParagraph.ChildObjects.RemoveAt(rangeIndex)
#Remove the objects which are behind the text range
tempList = []
for i in range(rangeIndex, ownParagraph.ChildObjects.Count):
    #Add a copy of these objects into a temp list
    tempP = ownParagraph.ChildObjects.get_Item(rangeIndex).Clone()
    tempList.append(tempP)
    ownParagraph.ChildObjects.RemoveAt(rangeIndex)
#Append field to the paragraph
ownParagraph.AppendField("MyFieldName", FieldType.FieldMergeField)
#Put these objects back into the paragraph one by one
for obj in tempList:
    ownParagraph.ChildObjects.Add(obj)

#Save doc file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()