from spire.doc import *
from spire.doc.common import *

# Input file path for the Word document containing the chart
input_file = "./Data/WordWithChart.docx"
out_file = "FormatDataLable.docx"

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

    # Iterate through each series in the chart
    for i in range(chart.Series.Count):
        # Get the current series
        series = chart.Series.get_Item(i)
        # Access the data labels for the series
        data_labels = series.DataLabels
        
        # Enable data labels for the series
        series.HasDataLabels = True
        # Show the data value on the labels
        data_labels.ShowValue = True
        # Hide the category name
        data_labels.ShowCategoryName = False
        # Hide the series name
        data_labels.ShowSeriesName = False
        # Disable leader lines
        data_labels.ShowLeaderLines = False
        # Set the separator between multiple data label values
        data_labels.Separator = ","
        # Apply a general number format (no specific formatting)
        data_labels.NumberFormat.FormatCode = "General"

        # Access the character formatting properties of the data labels
        fmt = data_labels.CharacterFormat
        # Set font size to 12 points
        fmt.FontSize = 12
        # Apply bold styling
        fmt.Bold = True
        # Disable strikethrough text effect
        fmt.IsStrikeout = False
        # Set text color to blue
        fmt.TextColor = Color.get_Blue()
        # Set border color to blue (if borders are used)
        fmt.Border.Color = Color.get_Blue()
        # Apply a low opacity effect (10% opacity)
        fmt.TextEffectFormat.TextOpacity = 0.1

# Save the modified document to a new file in DOCX format
doc.SaveToFile(out_file, FileFormat.Docx)

