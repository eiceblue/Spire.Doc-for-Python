from spire.doc import *
from spire.doc.common import *

inputFile1 = "./Data/Text2.docx"
inputFile2 = "./Data/Text1.docx"
outputFile = "ReplaceWithDocument.docx"

#Load a template document 
doc = Document(inputFile1)

#Load another document to replace text
replaceDoc = Document(inputFile2)
#Replace specified text with the other document
doc.Replace("Document1", replaceDoc, False, True)

#Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()