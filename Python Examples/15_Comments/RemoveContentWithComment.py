from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Comments.docx"
outputFile = "RemoveContentWithComment.docx"
#Create a document
document = Document()

#Load the document from disk.
document.LoadFromFile(inputFile)

#Get the first comment
comment = document.Comments[0]

#Get the paragraph of obtained comment
para = comment.OwnerParagraph

#Get index of the CommentMarkStart 
startIndex = para.ChildObjects.IndexOf(comment.CommentMarkStart)

#Get index of the CommentMarkEnd
endIndex = para.ChildObjects.IndexOf(comment.CommentMarkEnd)

#Create a list
dataList = []

#Get TextRanges between the indexes
for i in range(startIndex, endIndex):
    if isinstance(para.ChildObjects[i], TextRange):
        dataList.append( para.ChildObjects[i] if isinstance(para.ChildObjects[i], TextRange) else None)

#Insert a new TextRange
textRange = TextRange(document)

#Set text is null
textRange.Text = None

#Insert the new textRange
para.ChildObjects.Insert(endIndex, textRange)

#Remove previous TextRanges
for i, unusedItem in enumerate(dataList):
    para.ChildObjects.Remove(dataList[i])

#Save the document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
