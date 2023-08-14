from spire.doc.common import *
from spire.doc import *

outputFile = "HelloWorld.docx"
#Create a word document
document = Document()

#Create a new section
section = document.AddSection()

#Create a new paragraph
paragraph = section.AddParagraph()

#Append Text
paragraph.AppendText("Hello World!")

#Save doc file.
document.SaveToFile(outputFile, FileFormat.Docx)

#Close the document object
document.Close()

