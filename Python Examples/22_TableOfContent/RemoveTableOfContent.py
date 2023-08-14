import re
from spire.doc import *
from spire.doc.common import *

outputFile = "RemoveTableOfContent.docx"
inputFile = "./Data/TableOfContent.docx"

#Create a document
document = Document()

#Load the document from disk.
document.LoadFromFile(inputFile)

#Get the first body from the first section
body = document.Sections[0].Body

#Remove TOC from first body
regexStr = "TOC\\w+"
i = 0
while i < body.Paragraphs.Count:
    if re.match(regexStr,body.Paragraphs[i].StyleName):
        body.Paragraphs.RemoveAt(i)
        i -= 1
    i += 1
#Save the document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
