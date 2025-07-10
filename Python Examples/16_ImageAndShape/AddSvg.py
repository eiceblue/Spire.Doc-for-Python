from spire.doc import *
from spire.doc.common import *

inputFile = "Data/sample.svg"
outputFile = "AddSvg.docx"

#Create a Word document.
document = Document()

#Create a new section
section = document.AddSection()

#Add a new paragraph
para = section.AddParagraph()

#add a svg file to the paragraph
svgPicture = para.AppendPicture(inputFile)

#Set svg's width
svgPicture.Width = 200

#Set svg's height
svgPicture.Height = 200

#Save to file
document.SaveToFile(outputFile, FileFormat.Docx2016)

# Dispose the document
document.Dispose()
