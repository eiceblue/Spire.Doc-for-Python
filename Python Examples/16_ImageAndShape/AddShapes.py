import math
from spire.doc import *
from spire.doc.common import *

outputFile = "AddShapes.docx"

#Create Word document.
doc = Document()
sec = doc.AddSection()
para = sec.AddParagraph()
x = 60
y = 40
lineCount = 0
for i in range(1, 20):
    if lineCount > 0 and math.fmod(lineCount, 8) == 0:
        para.AppendBreak(BreakType.PageBreak)
        x = 60
        y = 40
        lineCount = 0
    #Add shape and set its size and position.
    shape = para.AppendShape(50, 50, ShapeType(i))
    shape.HorizontalOrigin = HorizontalOrigin.Page
    shape.HorizontalPosition = x
    shape.VerticalOrigin = VerticalOrigin.Page
    shape.VerticalPosition = y + 50
    x = x + int(shape.Width) + 50
    if i > 0 and math.fmod(i, 5) == 0:
        y = y + int(shape.Height) + 120
        lineCount += 1
        x = 60

doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
