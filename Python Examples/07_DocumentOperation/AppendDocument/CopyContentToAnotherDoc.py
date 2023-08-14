from spire.doc import *
from spire.doc.common import *

inputFile1 = "./Data/Sample.docx"
inputFile2 = "./Data/SampleB_1.docx"
outputFile = "CopyContentToAnotherDoc.docx"

#Initialize a new object of Document class and load the source document.
sourceDoc = Document()
sourceDoc.LoadFromFile(inputFile1)
#Initialize another object to load target document.
destinationDoc = Document()
destinationDoc.LoadFromFile(inputFile2)
#Copy content from source file and insert them to the target file.
for i in range(sourceDoc.Sections.Count):
    sec = sourceDoc.Sections.get_Item(i)
    for j in range(sec.Body.ChildObjects.Count):
        obj = sec.Body.ChildObjects.get_Item(j)
        destinationDoc.Sections[0].Body.ChildObjects.Add(obj.Clone())     
#Save to file.
destinationDoc.SaveToFile(outputFile, FileFormat.Docx2013)
sourceDoc.Close()
destinationDoc.Close()