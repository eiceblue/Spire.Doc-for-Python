from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/BlankTemplate.docx"
outputFile = "InsertImage.docx"

# Load Document
doc = Document()
doc.LoadFromFile(inputFile)

section = doc.Sections[0]
paragraph = section.Paragraphs[0] if section.Paragraphs.Count > 0 else section.AddParagraph()
paragraph.AppendText("The sample demonstrates how to insert an image into a document.")
paragraph.ApplyStyle(BuiltinStyle.Heading2)
paragraph = section.AddParagraph()
paragraph.AppendText("The above is a picture.")

# Create a picture
picture = DocPicture(doc)
picture.LoadImage("./Data/Word.png")

# set image's position
picture.HorizontalPosition = 50.0
picture.VerticalPosition = 60.0

# set image's size
picture.Width = 200.0
picture.Height = 200.0

# set textWrappingStyle with image
picture.TextWrappingStyle = TextWrappingStyle.Through
# Insert the picture at the beginning of the second paragraph
paragraph.ChildObjects.Insert(0, picture)

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
