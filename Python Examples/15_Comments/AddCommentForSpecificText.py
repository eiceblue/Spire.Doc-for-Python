ddef _InsertComments(doc, keystring):
    #Find the key string
    find = doc.FindString(keystring, False, True)

    #Create the commentmarkStart and commentmarkEnd
    commentmarkStart = CommentMark(doc)
    commentmarkStart.Type = CommentMarkType.CommentStart
    commentmarkStart.CommentId = 1
    commentmarkEnd = CommentMark(doc)
    commentmarkEnd.CommentId = 1
    commentmarkEnd.Type = CommentMarkType.CommentEnd

    #Add the content for comment
    comment = Comment(doc)
    comment.Format.CommentId = 1
    comment.Body.AddParagraph().Text = "Test comments"
    comment.Format.Author = "E-iceblue"

    #Get the textRanges
    range = find.GetRanges()
    length = len(range)
 
    #Get its paragraph
    para = range[0].OwnerParagraph

    #Get the index of textRange 
    index = para.ChildObjects.IndexOf(range[0])
    print(index)
    #Insert the commentmarkStart and commentmarkEnd
    para.ChildObjects.Insert(index, commentmarkStart)
    para.ChildObjects.Insert(index + length+1, commentmarkEnd)
    para.ChildObjects.Add(comment)


#Load the document from disk.
document = Document()
document.LoadFromFile(inputFile)

_InsertComments(document, "且仅")

#Save the document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()