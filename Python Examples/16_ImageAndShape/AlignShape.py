import unittest
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Shapes.docx"
outputFile = "AlignShape.docx"

#Load Document
doc = Document()
doc.LoadFromFile(inputFile)

section = doc.Sections[0]

for i in range(section.Paragraphs.Count):
    para = section.Paragraphs.get_Item(i)
    for j in range(para.ChildObjects.Count):
        obj = para.ChildObjects.get_Item(j)
        if isinstance(obj, ShapeObject):
            #Set the horizontal alignment as center
            ( obj if isinstance(obj, ShapeObject) else None).HorizontalAlignment = ShapeHorizontalAlignment.Center

            #//Set the vertical alignment as top
            #(obj as ShapeObject).VerticalAlignment = ShapeVerticalAlignment.Top

#Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()

