from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/InsertWordArt.docx"
outputFile = "InsertWordArt_out.docx"

# Create Word document.
doc = Document()

# Load Word document.
doc.LoadFromFile(inputFile)

# Add a paragraph.
paragraph = doc.Sections[0].AddParagraph()

# Add a shape.
shape = paragraph.AppendShape(250, 70, ShapeType.TextWave4)

# Set the position of the shape.
shape.VerticalPosition = 20
shape.HorizontalPosition = 80

# set the text of WordArt.
shape.WordArt.Text = "Thanks for reading."

# Set the fill color.
shape.FillColor = Color.get_Red()

# Set the border color of the text.
shape.StrokeColor = Color.get_Yellow()

# Save docx file.
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()
