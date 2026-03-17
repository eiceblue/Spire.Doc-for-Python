from spire.doc import *
from spire.doc.common import *

# Create a new Document instance
document = Document()

# Add a section to the document
section = document.AddSection()

# Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Column chart.")

# Add a new paragraph to the section
newParagraph = section.AddParagraph()

# Append a scatter chart shape to the paragraph with specified width and height
shape = newParagraph.AppendChart(ChartType.Column, float(500), float(300))

# Get the chart object from the shape
chart = shape.Chart

# Clear any existing series in the chart
chart.Series.Clear()

# Set the title of the chart
chart.Title.Text = "My chart"
        
# Add multiple series to the chart with categories (X-axis values) and corresponding data values
chart.Series.Add("Series 1", ["Word", "PDF", "Excel", "GoogleDocs", "Office"],
                 [190.0, 550.0, 210.0, 420.0, 150.0])

chart.Series.Add("Series 2", ["Word", "PDF", "Excel", "GoogleDocs", "Office"],
                 [400.0, 500.0, 110.0, 400.0, 250.0])

chart.Series.Add("Series 3", ["Word", "PDF", "Excel", "GoogleDocs", "Office"],
                 [550.0, 820.0, 180.0, 380.0, 130.0])
        
# Define the minimum value for the Y-axis as a float
data1 = 100.0

# Create an AxisBound object using the minimum value
axisBound_Minimum = AxisBound(data1)

# Set the minimum bound of the chart's Y-axis to the created AxisBound object
chart.AxisY.Bounds.Minimum = axisBound_Minimum

# Define the maximum value for the Y-axis as a float
data2 = 600.0

# Create an AxisBound object using the maximum value
axisBound_Maximum = AxisBound(data2)

# Set the maximum bound of the chart's Y-axis to the created AxisBound object
chart.AxisY.Bounds.Maximum = axisBound_Maximum

# Set the minor tick interval on the X-axis to 10 units
chart.AxisX.Units.Minor = 10

# Set the major tick interval on the X-axis to 100 units
chart.AxisX.Units.Major = 100

# Save the document to a file in Docx format
document.SaveToFile("SetBoundsOfChartAxes.docx", FileFormat.Docx)

document.Close()

# Dispose of the document object when finished using it
document.Dispose()