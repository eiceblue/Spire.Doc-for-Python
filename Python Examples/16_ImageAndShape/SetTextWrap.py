from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/ImageTemplate.docx"
outputFile = "SetTextWrap.docx"

# Load Document
doc = Document()
doc.LoadFromFile(inputFile)

for i in range(doc.Sections.Count):
    sec = doc.Sections.get_Item(i)
    for j in range(sec.Paragraphs.Count):
        para = sec.Paragraphs.get_Item(j)
        pictures = []
        # Get all pictures in the Word document
        for k in range(para.ChildObjects.Count):
            docObj = para.ChildObjects.get_Item(k)
            if docObj.DocumentObjectType == DocumentObjectType.Picture:
                pictures.append(docObj)

        # Set text wrap styles for each piture
        for pic in pictures:
            picture = pic if isinstance(pic, DocPicture) else None
            picture.TextWrappingStyle = TextWrappingStyle.Through
            picture.TextWrappingType = TextWrappingType.Both

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
