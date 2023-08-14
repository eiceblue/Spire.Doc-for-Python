from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/CommentTemplate.docx"
outputFile = "Comment.docx"

def _InsertComments( section):
    #Insert comment.
    paragraph = section.Paragraphs[1]
    comment = paragraph.AppendComment("Spire.Doc for .NET")
    comment.Format.Author = "E-iceblue"
    comment.Format.Initial = "CM"


#Load the document from disk.
document = Document()
document.LoadFromFile(inputFile)

_InsertComments(document.Sections[0])

#Save the document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()


