from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/CommentTemplate.docx"
outputFile = "InsertPictureIntoComment.docx"


#Load the document
doc = Document()
doc.LoadFromFile(inputFile)

#Get the first paragraph and insert comment
paragraph = doc.Sections[0].Paragraphs[2]
comment = paragraph.AppendComment("This is a comment.")
comment.Format.Author = "E-iceblue"

#Load a picture
docPicture = DocPicture(doc)
docPicture.LoadImage("./Data/E-iceblue.png")
#Insert the picture into the comment body
comment.Body.AddParagraph().ChildObjects.Add(docPicture)

#Save and launch
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()

