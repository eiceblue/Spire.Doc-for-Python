from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Shapes.docx"
outputFile = "ResetShapeSize.docx"

# Load Document
doc = Document()
doc.LoadFromFile(inputFile)

# Get the first section and the first paragraph that contains the shape
section = doc.Sections[0]
para = section.Paragraphs[0]

# Get the second shape and reset the width and height for the shape
shape = para.ChildObjects[1] if isinstance(
    para.ChildObjects[1], ShapeObject) else None
shape.Width = 200
shape.Height = 200

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
