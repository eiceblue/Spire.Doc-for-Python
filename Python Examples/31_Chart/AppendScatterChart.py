from spire.doc import *
from spire.doc.common import *
from spire.doc.charts import ChartType

outputFile = "AppendScatterChart.docx"

# Create a new instance of Document
document = Document()

# Add a section to the document
section = document.AddSection()

# Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Scatter chart.")

# Add a new paragraph to the section
newPara = section.AddParagraph()

# Append a scatter chart shape to the paragraph with specified width and height
shape = newPara.AppendChart(ChartType.Scatter, 450.0, 300.0)
chart = shape.Chart

# Clear any existing series in the chart
chart.Series.Clear()

# Add a new series to the chart with data points for X and Y values
chart.Series.Add("Scatter chart", [1.0, 2.0, 3.0, 4.0, 5.0],
                 [1.0, 20.0, 40.0, 80.0, 160.0])

# Save the document to a file in Docx format
document.SaveToFile(outputFile, FileFormat.Docx)

# Dispose of the document object when finished using it
document.Dispose()
