from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/CommentSample.docx"
outputFile = "RemoveAndReplaceComment.docx"
#Load the document
doc = Document()
doc.LoadFromFile(inputFile)

#Replace the content of the first comment
doc.Comments[0].Body.Paragraphs[0].Replace("This is the title", "This comment is changed.", False, False)

#Remove the second comment
doc.Comments.RemoveAt(1)

#Save and launch
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()

