from spire.doc import *
from spire.doc.common import *

outputFile = "RemoveImageWatermark.docx"
inputFile = "./Data/RemoveImageWatermark.docx"

#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)

#Set the watermark as null to remove the text and image watermark.
document.Watermark = None

#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()