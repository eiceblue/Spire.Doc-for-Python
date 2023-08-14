from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/ImageTemplate.docx"
outputFile = "ReplaceImageWithText_out.docx"

# Load Document
doc = Document()
doc.LoadFromFile(inputFile)

# Replace all pictures with texts
j = 1
for k in range(doc.Sections.Count):
    sec = doc.Sections.get_Item(k)
    for m in range(sec.Paragraphs.Count):
        para = sec.Paragraphs.get_Item(m)
        pictures = []
        # Get all pictures in the Word document
        for x in range(para.ChildObjects.Count):
            docObj = para.ChildObjects.get_Item(x)
            if docObj.DocumentObjectType == DocumentObjectType.Picture:
                pictures.append(docObj)

        # Replace pitures with the text "Here was image {image index}"
        for pic in pictures:
            index = para.ChildObjects.IndexOf(pic)
            textRange = TextRange(doc)
            textRange.Text = "Here was image {0}".format(j)
            para.ChildObjects.Insert(index, textRange)
            para.ChildObjects.Remove(pic)
            j += 1

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
