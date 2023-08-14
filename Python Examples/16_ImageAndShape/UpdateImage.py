from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/ImageTemplate.docx"
outputFile = "UpdateImage.docx"

# Load Document
doc = Document()
doc.LoadFromFile(inputFile)

# Get all pictures in the Word document
pictures = []
for i in range(doc.Sections.Count):
    sec = doc.Sections.get_Item(i)
    for j in range(sec.Paragraphs.Count):
        para = sec.Paragraphs.get_Item(j)
        for k in range(para.ChildObjects.Count):
            docObj = para.ChildObjects.get_Item(k)
            if docObj.DocumentObjectType == DocumentObjectType.Picture:
                pictures.append(docObj)

# Replace the first picture with a new image file
picture = pictures[0] if isinstance(pictures[0], DocPicture) else None
picture.LoadImage("./Data/E-iceblue.png")

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
