from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/CommentTemplate.docx"
outputFile = "AddCommentForSpecificText.docx"


def _InsertComments(doc, keystring):
    #Find the key string
    find = doc.FindString(keystring, False, True)

    #Create the commentmarkStart and commentmarkEnd
    commentmarkStart = CommentMark(doc)
    commentmarkStart.Type = CommentMarkType.CommentStart
    commentmarkEnd = CommentMark(doc)
    commentmarkEnd.Type = CommentMarkType.CommentEnd

    #Add the content for comment
    comment = Comment(doc)
    comment.Body.AddParagraph().Text = "Test comments"
    comment.Format.Author = "E-iceblue"

    #Get the textRange
    range = find.GetAsOneRange()

    #Get its paragraph
    para = range.OwnerParagraph

    #Get the index of textRange 
    index = para.ChildObjects.IndexOf(range)

    #Add comment
    para.ChildObjects.Add(comment)

    #Insert the commentmarkStart and commentmarkEnd
    para.ChildObjects.Insert(index, commentmarkStart)
    para.ChildObjects.Insert(index + 2, commentmarkEnd)


#Load the document from disk.
document = Document()
document.LoadFromFile(inputFile)

_InsertComments(document, "development")

#Save the document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()


    
