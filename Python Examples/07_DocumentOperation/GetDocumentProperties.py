from spire.doc import *
from spire.doc.common import *


def WriteAllText(fname: str, text: List[str]):
    fp = open(fname, "w")
    for s in text:
        fp.write(s)
    fp.close()


inputFile = "data/Properties.docx"
outputFile = "GetDocumentProperties.txt"

#Create a document
document = Document()

#Load the document from disk.
document.LoadFromFile(inputFile)

#Create content to save
content = ""

#Get Builtin document properties
title = document.BuiltinDocumentProperties.Title
comments = document.BuiltinDocumentProperties.Comments
author = document.BuiltinDocumentProperties.Author
keywords = document.BuiltinDocumentProperties.Keywords
company = document.BuiltinDocumentProperties.Company

#Set string format for displaying
result = "The Builtin document properties:\r\nTitle: " + title + ".\r\nComments: " + comments + ".\r\nAuthor: " + author + ".\r\nKeywords: " + keywords + ".\r\nCompany: " + company

#Add result string to content
#content.AppendLine(result + "\r\nThe custom document properties:")
content += result
content += "\r\nThe custom document properties:"

#Get custom document properties
for i in range(document.CustomDocumentProperties.Count):
    customProperties = document.CustomDocumentProperties[
        i].Name + ": " + document.CustomDocumentProperties.get_Item(
            i).ToString()
    content += customProperties

#Save them to a txt file
WriteAllText(outputFile, str(content))
document.Close()
