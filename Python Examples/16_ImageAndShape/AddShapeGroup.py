from spire.doc import *
from spire.doc.common import *


outputFile = "AddShapeGroup.docx"

#create a document
doc = Document()
sec = doc.AddSection()

#add a new paragraph
para = sec.AddParagraph()
#add a shape group with the height and width
shapegroup = para.AppendShapeGroup(375, 462)
shapegroup.HorizontalPosition = 180
#calcuate the scale ratio
X = float((shapegroup.Width / 1000.0))
Y = float((shapegroup.Height / 1000.0))

txtBox = TextBox(doc)
txtBox.SetShapeType(ShapeType.RoundRectangle)
txtBox.Width = 125 / X
txtBox.Height = 54 / Y
paragraph = txtBox.Body.AddParagraph()
paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
paragraph.AppendText("Start")
txtBox.HorizontalPosition = 19 / X
txtBox.VerticalPosition = 27 / Y
txtBox.Format.LineColor = Color.get_Green()
shapegroup.ChildObjects.Add(txtBox)

arrowLineShape = ShapeObject(doc, ShapeType.DownArrow)
arrowLineShape.Width = 16 / X
arrowLineShape.Height = 40 / Y
arrowLineShape.HorizontalPosition = 69 / X
arrowLineShape.VerticalPosition = 87 / Y
arrowLineShape.StrokeColor = Color.get_Purple()
shapegroup.ChildObjects.Add(arrowLineShape)

txtBox = TextBox(doc)
txtBox.SetShapeType(ShapeType.Rectangle)
txtBox.Width = 125 / X
txtBox.Height = 54 / Y
paragraph = txtBox.Body.AddParagraph()
paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
paragraph.AppendText("Step 1")
txtBox.HorizontalPosition = 19 / X
txtBox.VerticalPosition = 131 / Y
txtBox.Format.LineColor = Color.get_Blue()
shapegroup.ChildObjects.Add(txtBox)

arrowLineShape = ShapeObject(doc, ShapeType.DownArrow)
arrowLineShape.Width = 16 / X
arrowLineShape.Height = 40 / Y
arrowLineShape.HorizontalPosition = 69 / X
arrowLineShape.VerticalPosition = 192 / Y
arrowLineShape.StrokeColor = Color.get_Purple()
shapegroup.ChildObjects.Add(arrowLineShape)

txtBox = TextBox(doc)
txtBox.SetShapeType(ShapeType.Parallelogram)
txtBox.Width = 149 / X
txtBox.Height = 59 / Y
paragraph = txtBox.Body.AddParagraph()
paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
paragraph.AppendText("Step 2")
txtBox.HorizontalPosition = 7 / X
txtBox.VerticalPosition = 236 / Y
txtBox.Format.LineColor = Color.get_BlueViolet()
shapegroup.ChildObjects.Add(txtBox)

arrowLineShape = ShapeObject(doc, ShapeType.DownArrow)
arrowLineShape.Width = 16 / X
arrowLineShape.Height = 40 / Y
arrowLineShape.HorizontalPosition = 66 / X
arrowLineShape.VerticalPosition = 300 / Y
arrowLineShape.StrokeColor = Color.get_Purple()
shapegroup.ChildObjects.Add(arrowLineShape)

txtBox = TextBox(doc)
txtBox.SetShapeType(ShapeType.Rectangle)
txtBox.Width = 125 / X
txtBox.Height = 54 / Y
paragraph = txtBox.Body.AddParagraph()
paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
paragraph.AppendText("Step 3")
txtBox.HorizontalPosition = 19 / X
txtBox.VerticalPosition = 345 / Y
txtBox.Format.LineColor = Color.get_Blue()
shapegroup.ChildObjects.Add(txtBox)

#save the document
doc.SaveToFile(outputFile, FileFormat.Docx2010)
doc.Close()
