from io import FileIO
from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Template_N5.docx"
outputFile = "GetTextByStyleName.txt"
#Load document from disk
doc = Document()
doc.LoadFromFile(inputFile)

#Create string builder
builder = ""

#Loop through sections
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    #Loop through paragraphs
    for j in range(section.Paragraphs.Count):
        para = section.Paragraphs.get_Item(j)
        #Find the paragraph whose style name is "Heading1"
        if para.StyleName == "Heading1":
            #Write the text of paragraph
            builder += para.Text
            builder += "\n"

#Write the contents in a TXT file
with FileIO(outputFile, mode="w") as f:
    f.write(builder.encode("utf-8"))
