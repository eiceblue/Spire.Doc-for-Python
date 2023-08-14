from io import FileIO
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ShapeWithAlternativeText.docx"
outputFile = "GetAlternativeText.docx"

#Create a document
document = Document()
#Create string builder
builder = ''
document.LoadFromFile(inputFile)

#Loop through shapes and get the AlternativeText
for i in range(document.Sections.Count):
    section = document.Sections.get_Item(i)
    for j in range(section.Paragraphs.Count):
        para = section.Paragraphs.get_Item(j)
        for k in range(para.ChildObjects.Count):
            obj = para.ChildObjects.get_Item(k)
            if isinstance(obj, ShapeObject):
                text = ( obj if isinstance(obj, ShapeObject) else None).AlternativeText
                #Append the alternative text in builder
                builder += text
                builder += '\n'

#Save doc file.
#Write the contents in a TXT file
with FileIO(outputFile, mode="w") as f:
    f.write(builder.encode("utf-8"))

document.Close()
