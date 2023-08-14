from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Template_Docx_1.docx"
outputFile = "RemoveAllParagraphs.docx"
    
#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)

#Remove paragraphs from the body of every section in the document
for i in range(document.Sections.Count):
    section = document.Sections.get_Item(i)
    section.Paragraphs.Clear()

#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()