from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/ImageTemplate.docx"
outputFile = "SetTransparentColorForImage.docx"

# Load Document
doc = Document()
doc.LoadFromFile(inputFile)

# Get the first paragraph in the first section
paragraph = doc.Sections[0].Paragraphs[0]

# Set the blue color of the image(s) in the paragraph to transperant
for k in range(paragraph.ChildObjects.Count):
    obj = paragraph.ChildObjects.get_Item(k)
    if isinstance(obj, DocPicture):
        picture = obj if isinstance(obj, DocPicture) else None
        picture.TransparentColor = Color.get_Blue()

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
