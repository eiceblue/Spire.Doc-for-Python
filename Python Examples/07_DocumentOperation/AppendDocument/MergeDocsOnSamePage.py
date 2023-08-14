from spire.doc import *
from spire.doc.common import *

inputFile1 = "./Data/Insert.docx"
inputFile2 = "./Data/TableOfContent.docx"
outputFile = "MergeDocsOnSamePage.docx"

#Create a document
document = Document()
#Load the source document from disk.
document.LoadFromFile(inputFile1)
#Clone a destination  document
destinationDocument = Document()
#Load the destination document from disk.
destinationDocument.LoadFromFile(inputFile2)
#Traverse sections
for i in range(document.Sections.Count):
    section = document.Sections.get_Item(i)
    #Traverse body ChildObjects
    for j in range(section.Body.ChildObjects.Count):
        obj = section.Body.ChildObjects.get_Item(j)
        #Clone to destination document at the same page
        destinationDocument.Sections[0].Body.ChildObjects.Add(obj.Clone())
#Save the document.
destinationDocument.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
destinationDocument.Close()