from spire.doc import *
from spire.doc.common import *

# Path to the input Word document containing the chart
input_file = "./Data/WordWithChart.docx"
# Path for the output Word document
outputFile = "AddLengendForChart.docx"

# Create a new Document object
doc = Document()
# Load the existing Word document from the specified file path
doc.LoadFromFile(input_file)

# Access the first paragraph in the first section of the document
para = doc.Sections[0].Paragraphs[0]
# Get the first child object within the paragraph
obj = para.ChildObjects.get_Item(0)

# Check if the child object is a Shape (which may contain a chart)
if isinstance(obj, ShapeObject):
    # Cast the object to a ShapeObject type
    shape = obj  
    # Access the chart embedded within the shape
    chart = shape.Chart 
    legend = chart.Legend
    legend.Show = True
    legend.Position = LegendPosition.Right
    legend.Overlay = False
    legend.CharacterFormat.FontSize = 9
    legend.CharacterFormat.TextColor = Color.get_Blue()
    legend.CharacterFormat.Italic = True 
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Dispose()