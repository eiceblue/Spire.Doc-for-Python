from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Template_N3.docx"
outputFile = "RemoveSectionContent.docx"
       

#Load the Word document from disk
doc = Document()
doc.LoadFromFile(inputFile)

#Loop through all sections
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    #Remove header content
    section.HeadersFooters.Header.ChildObjects.Clear()
    #Remove body content
    section.Body.ChildObjects.Clear()
    #Remove footer content
    section.HeadersFooters.Footer.ChildObjects.Clear()

#Save the Word document
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()

