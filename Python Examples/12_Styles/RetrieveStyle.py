from io import FileIO
from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Styles.docx"
outputFile = "RetrieveStyle.txt"
#Load a template document 
doc = Document()
doc.LoadFromFile(inputFile)

#Traverse all paragraphs in the document and get their style names through StyleName property
styleName = ''
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    for j in range(section.Paragraphs.Count):
        paragraph = section.Paragraphs.get_Item(j)
        styleName += paragraph.StyleName + "\r\n"

#Write the contents in a TXT file
with FileIO(outputFile, mode="w") as f:
    f.write(styleName.encode("utf-8"))
