from spire.doc import *
from spire.doc.common import *


outputFile = "CreateImageHyperlink.docx"
inputFile = "./Data/BlankTemplate.docx"
inputFile_1 = "./Data/Spire.Doc.png"

# Load Document
doc = Document()
doc.LoadFromFile(inputFile)

section = doc.Sections[0]
# Add a paragraph
paragraph = section.AddParagraph()
# Load an image to a DocPicture object
picture = DocPicture(doc)
# Add an image hyperlink to the paragraph
picture.LoadImage(inputFile_1)

paragraph.AppendHyperlink(
    "https://www.e-iceblue.com/Introduce/doc-for-python.html", picture, HyperlinkType.WebLink)

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
