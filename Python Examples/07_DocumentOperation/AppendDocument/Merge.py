from spire.doc import *
from spire.doc.common import *

inputFile1 = "./Data/Sample.docx"
inputFile2 = "./Data/SampleB_1.docx"
outputFile = "merge.docx"

#Create word document
document = Document()
document.LoadFromFile(inputFile1, FileFormat.Docx)
#Load second file 
documentMerge = Document()
documentMerge.LoadFromFile(inputFile2, FileFormat.Docx)
#merge
for i in range(documentMerge.Sections.Count):
    sec = documentMerge.Sections.get_Item(i)
    document.Sections.Add(sec.Clone())  
#Save as docx file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
documentMerge.Close()