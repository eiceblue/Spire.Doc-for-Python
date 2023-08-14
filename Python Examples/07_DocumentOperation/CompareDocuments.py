from spire.doc import *
from spire.doc.common import *

inputFile1 = "./Data/SupportDocumentCompare1.docx"
inputFile2 = "./Data/SupportDocumentCompare2.docx"
outputFile = "CompareDocuments.docx"

#Load the first document
doc1 = Document()
doc1.LoadFromFile(inputFile1)
#Load the second document
doc2 = Document()
doc2.LoadFromFile(inputFile2)
#Compare the two documents
doc1.Compare(doc2, "E-iceblue")
#Save as docx file.
doc1.SaveToFile(outputFile, FileFormat.Docx2013)
doc1.Close()
doc2.Close()