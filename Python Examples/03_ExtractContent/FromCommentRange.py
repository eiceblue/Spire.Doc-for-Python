from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Comments.docx"
outputFile = "FromCommentRange.docx"
       
#Create a document
sourceDoc = Document()

#Load the document from disk.
sourceDoc.LoadFromFile(inputFile)

#Create a destination document
destinationDoc = Document()

#Add section for destination document
destinationSec = destinationDoc.AddSection()

#Get the first comment
comment = sourceDoc.Comments[0]

#Get the paragraph of obtained comment
para = comment.OwnerParagraph

#Get index of the CommentMarkStart 
startIndex = para.ChildObjects.IndexOf(comment.CommentMarkStart)

#Get index of the CommentMarkEnd
endIndex = para.ChildObjects.IndexOf(comment.CommentMarkEnd)

paragraph = destinationSec.AddParagraph()

#Traverse paragraph ChildObjects
for i in range(startIndex, endIndex + 1):
    #Clone the ChildObjects of source document
    doobj = para.ChildObjects[i].Clone()

    #Add to destination document 
    paragraph.ChildObjects.Add(doobj)

#Save the destination document
destinationDoc.SaveToFile(outputFile, FileFormat.Docx)
destinationDoc.Close()
