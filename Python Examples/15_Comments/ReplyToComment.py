from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Comment.docx"
outputFile = "ReplyToComment.docx"

#Load the document from disk.
doc = Document()
doc.LoadFromFile(inputFile)

#get the first comment.
comment1 = doc.Comments[0]

#create a new comment and specify the author and content.
replyComment1 = Comment(doc)
replyComment1.Format.Author = "E-iceblue"
replyComment1.Body.AddParagraph().AppendText("Spire.Doc is a professional Word .NET library on operating Word documents.")

#add the new comment as a reply to the selected comment.
comment1.ReplyToComment(replyComment1)

docPicture = DocPicture(doc)
docPicture.LoadImage("./Data/logo.png")

#insert a picture in the comment
replyComment1.Body.Paragraphs[0].ChildObjects.Add(docPicture)

#Save the document.
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()

