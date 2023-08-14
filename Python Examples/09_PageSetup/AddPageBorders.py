from spire.doc import *
from spire.doc.common import *


inputFile =  "./Data/Template_Docx_1.docx"
outputFile = "AddPageBorders.docx"

#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)

#Add page borders with special style and color.
document.Sections[0].PageSetup.Borders.BorderType(BorderStyle.DoubleWave)
document.Sections[0].PageSetup.Borders.Color(Color.get_LightSeaGreen())

#Set the space between border and text.
document.Sections[0].PageSetup.Borders.Left.Space = 50
document.Sections[0].PageSetup.Borders.Right.Space  = 50


#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()