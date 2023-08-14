from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Sample.docx"
outputFile = "EnableTrackChanges.docx"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Enable the track changes.
document.TrackChanges = True
#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()