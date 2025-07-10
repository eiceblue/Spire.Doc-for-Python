from spire.doc import *
from spire.doc.common import *
from spire.doc.charts import ChartType

outputFile = "AppendPieChart.docx"

# Create a new instance of Document
document = Document()

# Add a section to the document
section = document.AddSection()

# Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Pie chart.")

# Add a new paragraph to the section
newPara = section.AddParagraph()

# Append a pie chart shape to the paragraph with specified width and height
shape = newPara.AppendChart(ChartType.Pie, 500.0, 300.0)
chart = shape.Chart

# Add a series to the chart with categories (labels) and corresponding data values
series = chart.Series.Add("Test Series", ["Word", "PDF", "Excel"],
                          [2.7, 3.2, 0.8])

# Save the document to a file in Docx format
document.SaveToFile(outputFile, FileFormat.Docx)

# Dispose of the document object when finished using it
document.Dispose()
