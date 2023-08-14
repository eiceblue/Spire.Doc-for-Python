from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Sample.docx"
outputFile = "SetWordViewModes.docx"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Set Word view modes.
document.ViewSetup.DocumentViewType = DocumentViewType.WebLayout
document.ViewSetup.ZoomPercent = 150
document.ViewSetup.ZoomType = ZoomType.none
#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()