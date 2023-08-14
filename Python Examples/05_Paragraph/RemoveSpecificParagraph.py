from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Template_Docx_1.docx"
outputFile = "RemoveSpecificParagraph.docx"
     
#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)
#Remove the first paragraph from the first section of the document.
document.Sections[0].Paragraphs.RemoveAt(0)

#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()