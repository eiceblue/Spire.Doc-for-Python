from spire.doc import *
from spire.doc.common import *

inputFile1 = "./Data/Sample.docx"
inputFile2 = "./Data/SampleB_1.docx"
outputFile = "SimpleInsertFile.docx"

#Load the Word document
doc = Document()
doc.LoadFromFile(inputFile1)
#Insert document from file
doc.InsertTextFromFile(inputFile2, FileFormat.Auto)
#Save the document
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()