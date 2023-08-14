from spire.doc import *
from spire.doc.common import *

def ExtractBetweenParagraphs(sourceDocument, destinationDocument, startPara, endPara):
        #Extract the content
        for i in range(startPara - 1, endPara):
            #Clone the ChildObjects of source document
            doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone()

            #Add to destination document 
            destinationDocument.Sections[0].Body.ChildObjects.Add(doobj)
        
inputFile = "./Data/Sample.docx"
outputFile = "BetweenParagraphs.docx"
    

#Create the first document
sourceDocument = Document()

#Load the source document from disk.
sourceDocument.LoadFromFile(inputFile)

#Create a destination document
destinationDoc = Document()

#Add a section
section = destinationDoc.AddSection()

#Extract content between the first paragraph to the third paragraph
ExtractBetweenParagraphs(sourceDocument, destinationDoc, 1, 3)

#Save the document.
destinationDoc.SaveToFile(outputFile, FileFormat.Docx)
destinationDoc.Close()
