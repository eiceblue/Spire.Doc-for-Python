from spire.doc import *
from spire.doc.common import *
from spire.doc.charts import ChartType

outputFile = "AppendBubbleChart.docx"

# Create a new instance of Document
document = Document()

# Add a section to the document
section = document.AddSection()

# Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Bubble chart.")

# Add a new paragraph to the section
newPara = section.AddParagraph()

# Append a bubble chart shape to the paragraph with specified width and height
shape = newPara.AppendChart(ChartType.Bubble, float(500), float(300))

# Get the chart object from the shape
chart = shape.Chart

# Clear any existing series in the chart
chart.Series.Clear()

# Add a new series to the chart with data points for X, Y, and bubble size values
series = chart.Series.Add("Test Series", [2.9, 3.5, 1.1, 4.0, 4.0],
                          [1.9, 8.5, 2.1, 6.0, 1.5], [9.0, 4.5, 2.5, 8.0, 5.0])

# Save the document to a file in Docx format
document.SaveToFile(outputFile, FileFormat.Docx)

# Dispose of the document object when finished using it
document.Dispose()
