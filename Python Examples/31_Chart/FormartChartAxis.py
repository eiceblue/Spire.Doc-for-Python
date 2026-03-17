from spire.doc import *
from spire.doc.common import *

# Path to the input Word document containing the chart
input_file = "./Data/WordWithChart.docx"
# Path for the output Word document
outputFile = "FormartChartAxis.docx"

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
    
    # Get all axes of the chart
    axes = chart.Axes
    
    # Iterate through all axes
    for index in range(len(axes)):
        axis = axes[index]
        
        # Process X-axis (first axis, usually the category axis)
        if index == 0:
            axis.CategoryType = AxisCategoryType.Category  # Set category type to Category
            axis.Bounds.Maximum = AxisBound(5.0)  # Set maximum boundary
            axis.Bounds.Minimum = AxisBound(0.0)  # Set minimum boundary
            axis.Units.Major = 1.0  # Set major unit
            axis.Units.MajorTimeUnit = AxisTimeUnit.Days  # Set major time unit
            axis.Units.Minor = 1.0  # Set minor unit
            axis.Units.MinorTimeUnit = AxisTimeUnit.Days  # Set minor time unit
            axis.HasMajorGridlines = False  # Disable major gridlines
            axis.HasMinorGridlines = True  # Enable minor gridlines
            axis.Labels.IsAutoSpacing = False  # Disable automatic label spacing
            axis.Labels.Spacing = 1  # Set label spacing
            axis.Labels.Offset = 1  # Set label offset
            axis.Labels.Position = AxisTickLabelPosition.Low  # Set label position to low
            axis.ReverseOrder = True  # Reverse axis order
            axis.Title.Text = "X-axis"  # Set axis title text
            axis.Title.Show = True  # Show axis title
            axis.Title.Overlay = True  # Allow title overlay
            
        # Process Y-axis (second axis, usually the value axis)
        elif index == 1:
            axis.CategoryType = AxisCategoryType.Automatic  # Set category type to Automatic
            axis.Units.IsMajorAuto = True  # Enable automatic major units
            axis.Units.IsMinorAuto = True  # Enable automatic minor units
            axis.Bounds.LogBase = 10  # Set logarithmic base to 10
            axis.HasMajorGridlines = True  # Enable major gridlines
            axis.HasMinorGridlines = False  # Disable minor gridlines
            axis.ReverseOrder = False  # Do not reverse axis order
            axis.Labels.IsAutoSpacing = True  # Enable automatic label spacing
            axis.Title.Text = "Y-axis"  # Set axis title text
            axis.Title.Show = True  # Show axis title
            axis.Title.Overlay = True  # Allow title overlay
            
        # Process Z-axis (third axis, usually for 3D charts)
        else:
            axis.Title.Text = "Z-axis"  # Set axis title text
            axis.Title.Show = True  # Show axis title
            axis.Title.Overlay = False  # Disable title overlay

        # Set common properties for all axes
        axis.Labels.Alignment = LabelAlignment.Left  # Set label alignment to left
        axis.Units.BaseTimeUnit = AxisTimeUnit.Days  # Set base time unit
        axis.AxisBetweenCategories = True  # Enable axis between categories
        axis.DisplayUnits.CustomUnit = 1  # Set custom display unit
        axis.DisplayUnits.Unit = AxisBuiltInUnit.Custom  # Use custom units
        axis.DisplayUnits.ShowLabel = True  # Show unit label
        axis.TickMarks.Spacing = 1  # Set tick mark spacing
        axis.TickMarks.Major = AxisTickMark.Cross  # Set major tick mark style to cross
        axis.TickMarks.Minor = AxisTickMark.Inside  # Set minor tick mark style to inside

        # Set format for axis title
        title_fmt = axis.Title.GetCharacterFormat()
        title_fmt.FontSize = 8  # Set font size
        title_fmt.TextColor = Color.get_Red()  # Set text color to red
        title_fmt.Bold = True  # Set bold

# Save the document to the output file path using Docx format
doc.SaveToFile(outputFile, FileFormat.Docx)

# Dispose the document object to release resources
doc.Dispose()