from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/TextInputField.docx"
outputFile = "StartFromFormField.docx"

#Create the source document
sourceDocument = Document()

#Load the source document from disk.
sourceDocument.LoadFromFile(inputFile)

#Create a destination document
destinationDoc = Document()

#Add a section
section = destinationDoc.AddSection()

#Define a variables
index = 0

#Traverse FormFields
for i in range(sourceDocument.Sections[0].Body.FormFields.Count):
    #Find FieldFormTextInput type field
    field = sourceDocument.Sections[0].Body.FormFields.get_Item(i)
    if field.Type == FieldType.FieldFormTextInput:
        #Get the paragraph
        paragraph = field.OwnerParagraph

        #Get the index
        index = sourceDocument.Sections[0].Body.ChildObjects.IndexOf(paragraph)
        break

#Extract the content
i = index
while i < index + 3:
    #Clone the ChildObjects of source document
    doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone()

    #Add to destination document 
    section.Body.ChildObjects.Add(doobj)
    i += 1

#Save the document.
destinationDoc.SaveToFile(outputFile, FileFormat.Docx)
destinationDoc.Close()

