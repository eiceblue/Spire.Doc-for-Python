from spire.doc import *
from spire.doc.common import *

def ExtractByTable(sourceDocument, destinationDocument, startPara, tableNo):
        #Get the table from the source document
        table = sourceDocument.Sections[0].Tables[tableNo - 1] if isinstance(sourceDocument.Sections[0].Tables[tableNo - 1], Table) else None

        #Get the table index
        index = sourceDocument.Sections[0].Body.ChildObjects.IndexOf(table)
        for i in range(startPara - 1, index + 1):
            #Clone the ChildObjects of source document
            doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone()

            #Add to destination document 
            destinationDocument.Sections[0].Body.ChildObjects.Add(doobj)


inputFile = "./Data/IncludingTable.docx"
outputFile = "FromParagraphToTable.docx"
      
#Create the first document
sourceDocument = Document()

#Load the source document from disk.
sourceDocument.LoadFromFile(inputFile)

#Create a destination document
destinationDoc = Document()

#Add a section
destinationSection = destinationDoc.AddSection()

#Extract the content from the first paragraph to the first table
ExtractByTable(sourceDocument, destinationDoc, 1, 1)

#Save the document.
destinationDoc.SaveToFile(outputFile, FileFormat.Docx)
destinationDoc.Close()