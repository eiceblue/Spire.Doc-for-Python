from spire.doc import *
from spire.doc.common import *

def ExtractBetweenParagraphStyles(sourceDocument, destinationDocument, stylename1, stylename2):
        startindex = 0
        endindex = 0
        #travel the sections of source document
        for i in range(sourceDocument.Sections.Count):
            section = sourceDocument.Sections.get_Item(i)
            #travel the paragraphs
            for j in range(section.Paragraphs.Count):
                paragraph = section.Paragraphs.get_Item(j)
                #Judge paragraph style1
                if paragraph.StyleName == stylename1:
                    #Get the paragraph index
                    startindex = section.Body.Paragraphs.IndexOf(paragraph)
                #Judge paragraph style2
                if paragraph.StyleName == stylename2:
                    #Get the paragraph index
                    endindex = section.Body.Paragraphs.IndexOf(paragraph)
            #Extract the content
            for k in range(startindex + 1, endindex):
                #Clone the ChildObjects of source document
                doobj = sourceDocument.Sections[0].Body.ChildObjects[k].Clone()

                #Add to destination document 
                destinationDocument.Sections[0].Body.ChildObjects.Add(doobj)

   
inputFile ="./Data/BetweenParagraphStyle.docx"
outputFile = "BetweenParagraphStyles.docx"

#Create the first document
sourceDocument = Document()

#Load the source document from disk.
sourceDocument.LoadFromFile(inputFile)

#Create a destination document
destinationDoc = Document()

#Add a section
section = destinationDoc.AddSection()

#Extract content between the first paragraph to the third paragraph
ExtractBetweenParagraphStyles(sourceDocument, destinationDoc, "1", "2")

#Save the document.
destinationDoc.SaveToFile(outputFile, FileFormat.Docx)
destinationDoc.Close()