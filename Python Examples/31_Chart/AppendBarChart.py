from spire.doc import *
from spire.doc.common import *
from spire.doc.charts import ChartType

outputFile = "AppendBarChart.docx"

# Create a new instance of Document
document = Document()

# Add a section to the document
section = document.AddSection()

# Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Bar chart.")

# Add a new paragraph to the section
newPara = section.AddParagraph()

# Append a bar chart shape to the paragraph with specified width and height
chartShape = newPara.AppendChart(ChartType.Bar, float(400), float(300))
chart = chartShape.Chart

# Get the title of the chart
title = chart.Title

# Set the text of the chart title
title.Text = "My Chart"

# Show the chart title
title.Show = True

# Overlay the chart title on top of the chart
title.Overlay = True

# Save the document to a file in Docx format
document.SaveToFile(outputFile, FileFormat.Docx)

# Dispose of the document object when finished using it
document.Dispose()
