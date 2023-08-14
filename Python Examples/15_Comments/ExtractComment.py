from io import FileIO
from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/CommentSample.docx"
outputFile = "ExtractComment.txt"

#Load the document
doc = Document()
doc.LoadFromFile(inputFile)

content = ''

#Traverse all comments
for i in range(doc.Comments.Count):
    comment = doc.Comments.get_Item(i)
    for j in range(comment.Body.Paragraphs.Count):
        p = comment.Body.Paragraphs.get_Item(j)
        content += p.Text
        content += '\n'

#Write the contents in a TXT file
with FileIO(outputFile, mode="w") as f:
    f.write(content.encode("utf-8"))

