from spire.doc import *
from spire.doc.common import *
from spire.doc.charts import ChartType

outputFile = "AppendColumnChart.docx"

# Create a new instance of Document
document = Document()

# Add a section to the document
section = document.AddSection()

# Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Column chart.")

# Add a new paragraph to the section
newPara = section.AddParagraph()

# Append a column chart shape to the paragraph with specified width and height
shape = newPara.AppendChart(ChartType.Column, float(500), float(300))

# Get the chart object from the shape
chart = shape.Chart

# Clear any existing series in the chart
chart.Series.Clear()

# Add a new series to the chart with data points for X values (categories) and Y values
chart.Series.Add("Test Series",
                 ["Word", "PDF", "Excel", "GoogleDocs", "Office"], [
                     float(1900000),
                     float(850000),
                     float(2100000),
                     float(600000),
                     float(1500000)
                 ])

# Set the number format for the Y-axis labels
chart.AxisY.NumberFormat.FormatCode = "#,##0"

# Save the document to a file in Docx format
document.SaveToFile(outputFile, FileFormat.Docx)

# Dispose of the document object when finished using it
document.Dispose()
