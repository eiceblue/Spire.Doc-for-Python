from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Sample.docx"
outputFile = "CloneWordDocument.docx"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Clone the word document.
newDoc = document.Clone()
#Save the file.
newDoc.SaveToFile(outputFile, FileFormat.Docx2013)
newDoc.Close()
document.Close()