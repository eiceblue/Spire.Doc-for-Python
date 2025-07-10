from spire.doc import *
from spire.doc.common import *
from spire.doc.charts import ChartType

outputFile = "AppendLineChart.docx"

# Create a new instance of Document
document = Document()

# Add a section to the document
section = document.AddSection()

# Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Line chart.")

# Add a new paragraph to the section
newPara = section.AddParagraph()

# Append a line chart shape to the paragraph with specified width and height
shape = newPara.AppendChart(ChartType.Line, 500.0, 300.0)

# Get the chart object from the shape
chart = shape.Chart

# Get the title of the chart
title = chart.Title

# Set the text of the chart title
title.Text = "My Chart"

# Clear any existing series in the chart
seriesColl = chart.Series
seriesColl.Clear()

# Define categories (X-axis values)
categories = ["C1", "C2", "C3", "C4", "C5", "C6"]

# Add two series to the chart with specified categories and Y-axis values
seriesColl.Add("AW Series 1", categories, [1.0, 2.0, 2.5, 4.0, 5.0, 6.0])
seriesColl.Add("AW Series 2", categories, [2.0, 3.0, 3.5, 6.0, 6.5, 7.0])

# Save the document to a file in Docx format
document.SaveToFile(outputFile, FileFormat.Docx)

# Dispose of the document object when finished using it
document.Dispose()
