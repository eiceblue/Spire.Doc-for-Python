from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_2.docx"
outputFile = "AddGutter.docx"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Create a new section
section = document.Sections[0]
#Set gutter
section.PageSetup.Gutter = 100
#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()