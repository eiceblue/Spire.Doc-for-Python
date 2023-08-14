from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ImageTemplate.docx"
outputFile = "ResetImageSize_out.docx"

# Load Document
doc = Document()
doc.LoadFromFile(inputFile)

# Get the first secion
section = doc.Sections[0]
# Get the first paragraph
paragraph = section.Paragraphs[0]

# Reset the image size of the first paragraph
for i in range(paragraph.ChildObjects.Count):
    docObj = paragraph.ChildObjects.get_Item(i)
    if isinstance(docObj, DocPicture):
        picture = DocPicture(docObj)
        picture.Width = 50
        picture.Height = 50

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
