from spire.doc import *
from spire.doc.common import *

outputFile = "DetectAndRemoveVBAMacros.docm"
inputFile = "./Data/DetectAndRemoveVBAMacros.docm"

#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)

#If the document contains Macros, remove them from the document.
if document.IsContainMacro:
    document.ClearMacros()
    
#Save to file.
document.SaveToFile(outputFile, FileFormat.Docm)
document.Close()
