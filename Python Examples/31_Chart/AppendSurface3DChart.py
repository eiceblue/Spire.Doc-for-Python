from spire.doc import *
from spire.doc.common import *
from spire.doc.charts import ChartType

outputFile = "AppendSurface3DChart.docx"

# Create a new instance of Document
document = Document()

# Add a section to the document
section = document.AddSection()

# Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Surface3D chart.")

# Add a new paragraph to the section
newPara = section.AddParagraph()

# Append a Surface3D chart shape to the paragraph with specified width and height
shape = newPara.AppendChart(ChartType.Surface3D, 500.0, 300.0)

# Get the chart object from the shape
chart = shape.Chart

# Clear any existing series in the chart
chart.Series.Clear()

# Set the title of the chart
chart.Title.Text = "My chart"

# Add multiple series to the chart with categories (X-axis values) and corresponding data values
chart.Series.Add("Series 1", ["Word", "PDF", "Excel", "GoogleDocs", "Office"],
                 [1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0])

chart.Series.Add("Series 2", ["Word", "PDF", "Excel", "GoogleDocs", "Office"],
                 [900000.0, 50000.0, 1100000.0, 400000.0, 2500000.0])

chart.Series.Add("Series 3", ["Word", "PDF", "Excel", "GoogleDocs", "Office"],
                 [500000.0, 820000.0, 1500000.0, 400000.0, 100000.0])

# Save the document to a file in Docx format
document.SaveToFile(outputFile, FileFormat.Docx)

# Dispose of the document object when finished using it
document.Dispose()
