from spire.doc import *
from spire.doc.common import *
from spire.doc.charts.ChartType import ChartType

outputFile = "AppendBarChart.docx"
# Create a new instance of Document
document = Document()

# Add a section to the document
section = document.AddSection()

# Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Bar chart.")

# Add a new paragraph to the section
newPara = section.AddParagraph()

# Insert a bar chart into the paragraph with width 400 and height 300
chartShape = newPara.AppendChart(ChartType.Bar, float(400), float(300))

# Get the chart object from the chart shape for further configuration
chart = chartShape.Chart

# Access the chart's title properties
title = chart.Title

# Enable display of the chart title
title.Show = True

# Set the title to overlay on top of the chart area
title.Overlay = True

# Set the text content of the title
title.Text = "My Chart"

# Access the character formatting options for the title text
fmt = title.CharacterFormat

# Set the font size to 12 points
fmt.FontSize = 12

# Apply bold styling to the title text
fmt.Bold = True

# Disable strikethrough effect on the text
fmt.IsStrikeout = False

# Set the text color to blue
fmt.TextColor = Color.get_Blue()

# Save the document with charts to a file named "AppendBarChart.docx"
document.SaveToFile(outputFile, FileFormat.Docx)

# Properly release the document object to free up system resources
document.Dispose()