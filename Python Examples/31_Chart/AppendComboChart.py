from spire.doc import *
from spire.doc.common import *
from spire.doc.charts.ChartType import ChartType

outputFile = "AppendComboChart.docx"
# Create a new Word document
doc = Document()

# Add a combo chart with column and line series
para1 = doc.AddSection().AddParagraph()
chart1 = para1.AppendChart(ChartType.Column, 450.0, 300.0).Chart
# Change Series 3 to line chart (secondary axis)
chart1.ChangeSeriesType("Series 3", ChartSeriesType.Line, True)  
# Updated chart title to English
chart1.Title.Text = "Column + Line Combo Chart"  

# Add a combo chart with column and area series
para2 = doc.AddSection().AddParagraph()
chart2 = para2.AppendChart(ChartType.Column, 450.0, 300.0).Chart
# Change Series 3 to area chart (primary axis)
chart2.ChangeSeriesType("Series 3", ChartSeriesType.Area, False)  
 # Updated chart title to English
chart2.Title.Text = "Column + Area Combo Chart" 

# Add a combo chart with line and pie series (note: pie charts typically require special handling)
para3 = doc.AddSection().AddParagraph()
chart3 = para3.AppendChart(ChartType.Line, 450.0, 300.0).Chart
# Change first series to pie chart
chart3.ChangeSeriesType(chart3.Series[0].Name, ChartSeriesType.Pie, True)  
# Updated chart title to English
chart3.Title.Text = "Line + Pie Combo Chart"  

# Add a combo chart with line and radar series
para4 = doc.AddSection().AddParagraph()
chart4 = para4.AppendChart(ChartType.Line, 450.0, 300.0).Chart
# Change Series 3 to radar chart
chart4.ChangeSeriesType("Series 3", ChartSeriesType.Radar, True)  
# Updated chart title to English
chart4.Title.Text = "Line + Radar Combo Chart"  

# Add a combo chart with area and scatter series
para5 = doc.AddSection().AddParagraph()
chart5 = para5.AppendChart(ChartType.Area, 450.0, 300.0).Chart
# Change first series to scatter chart
chart5.ChangeSeriesType(chart5.Series[0].Name, ChartSeriesType.Scatter, True)  
# Updated chart title to English
chart5.Title.Text = "Area + Scatter Combo Chart"  

# Add a line chart with different axis configurations for series
para6 = doc.AddSection().AddParagraph()
chart6 = para6.AppendChart(ChartType.Line, 450.0, 300.0).Chart
# Series 1 on primary axis
chart6.ChangeSeriesType("Series 1", ChartSeriesType.Line, False)  
# Series 2 on secondary axis
chart6.ChangeSeriesType("Series 2", ChartSeriesType.Line, True)  
# Updated chart title to English

chart6.Title.Text = "Dual Axis Line Chart"  
# Add a multi-type combo chart with column, line, and scatter series
para7 = doc.AddSection().AddParagraph()
chart7 = para7.AppendChart(ChartType.Column, 450.0, 300.0).Chart
# First series to line chart
chart7.ChangeSeriesType(chart7.Series[0].Name, ChartSeriesType.Line, False)  
# Second series to scatter chart
chart7.ChangeSeriesType(chart7.Series[1].Name, ChartSeriesType.Scatter, True) 


# Save the document to file
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Dispose()