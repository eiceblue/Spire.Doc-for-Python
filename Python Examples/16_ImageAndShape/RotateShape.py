from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Shapes.docx"
outputFile = "RotateShape.docx"

# Load Document
doc = Document()
doc.LoadFromFile(inputFile)

# Get the first section
section = doc.Sections[0]

# Traverse the word document and set the shape rotation as 20
for i in range(section.Paragraphs.Count):
    para = section.Paragraphs.get_Item(i)
    for j in range(para.ChildObjects.Count):
        obj = para.ChildObjects.get_Item(j)
        if isinstance(obj, ShapeObject):
            (obj if isinstance(obj, ShapeObject) else None).Rotation = 20.0

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
