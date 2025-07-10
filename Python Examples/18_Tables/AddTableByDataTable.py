from spire.doc import *
from spire.doc.common import *
import xml.etree.ElementTree as ET
import pandas as pd


@staticmethod
def _FillTableUsingDataTable(table, dataTable):
    columnCount = len(dataTable[0])

    for dataRow in dataTable:
        row = table.AddRow(columnCount)
        i = 0
        for col in dataRow:
            #columnIndex = dataTable.Columns.IndexOf(dataColumn)
            value = str(col.text)
            cell = row.Cells.get_Item(i)
            paragraph = cell.AddParagraph()
            paragraph.AppendText(value)
            #Set the alignment of cell
            cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
            paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
            i += 1


outputFile = "AddTableByDataTable.docx"
inputFile = "Data/dataTable.xml"

#Create a Word document
document = Document()

#Get the first section
section = document.AddSection()

#Add paragraph style
style = ParagraphStyle(document)
style.CharacterFormat.FontSize = 20
style.CharacterFormat.Bold = True
style.CharacterFormat.TextColor = Color.get_CadetBlue()
document.Styles.Add(style)

#Create a paragraph and append text
para = section.AddParagraph()
para.AppendText("Table")
#Apply style
para.Format.HorizontalAlignment = HorizontalAlignment.Center
para.ApplyStyle(style.Name)

#Load data
#ds = DataSet()
#ds.ReadXml(inputFile)
tree = ET.parse(inputFile)
dataTable = tree.getroot()

#Get the first data table
#dataTable = ds.Tables[0]

#Add a table
table = section.AddTable(True)
#Set its width
table.PreferredWidth = PreferredWidth(WidthType.Percentage, 100)

#Fill table with the data of datatable
_FillTableUsingDataTable(table, dataTable)

#Set table style
table.Format.Paddings.SetAll(5)
row = table.FirstRow
i = 0
while i < row.Cells.Count:
    row.Cells.get_Item(i).CellFormat.Shading.BackgroundPatternColor = Color.get_CadetBlue()
    i += 1

#Save the Word file
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
