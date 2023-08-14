from spire.doc import *
from spire.doc.common import *

inputFile1 = "./Data/Sample.docx"
inputFile2 = "./Data/SampleB_1.docx"
outputFile = "KeepSameFormat.docx"

#Load the source document from disk
srcDoc = Document()
srcDoc.LoadFromFile(inputFile1)
#Load the destination document from disk
destDoc = Document()
destDoc.LoadFromFile(inputFile2)
#Keep same format of source document
srcDoc.KeepSameFormat = True
#Copy the sections of source document to destination document
for i in range(srcDoc.Sections.Count):
    section = srcDoc.Sections.get_Item(i)
    destDoc.Sections.Add(section.Clone())
#Save the Word document
destDoc.SaveToFile(outputFile, FileFormat.Docx2013)
srcDoc.Close()
destDoc.Close()
