from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Sample.docx"
outputFile = "LoadAndSaveToDisk.docx"

#Create a new document
doc = Document()
# Load the document from the absolute/relative path on disk.
doc.LoadFromFile(inputFile)
# Save the document to disk
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
