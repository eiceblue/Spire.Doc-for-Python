
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_4.docx"
outputFile = "RemovePageBreaks.docx"
#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Traverse every paragraph of the first section of the document.
for j in range(document.Sections[0].Paragraphs.Count):
    p = document.Sections[0].Paragraphs[j]
    #Traverse every child object of a paragraph.
    for i in range(p.ChildObjects.Count):
        obj = p.ChildObjects[i]
        #Find the page break object.
        if obj.DocumentObjectType == DocumentObjectType.Break:
            b = obj if isinstance(obj, Break) else None
            #Remove the page break object from paragraph.
            p.ChildObjects.Remove(b)
#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
