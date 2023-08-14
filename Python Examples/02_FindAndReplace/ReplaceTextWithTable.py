from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_1.docx"
outputFile = "ReplaceTextWithTable.docx"

#Create Word document.
document=Document()

#Load the file from disk.
document.LoadFromFile(inputFile)

#Return TextSection by finding the key text string "ChristmasDay,December25".
section=document.Sections[0]
selection=document.FindString("Christmas Day, December 25",True,True)

#Return TextRange from TextSection, then get OwnerParagraph through TextRange.
range=selection.GetAsOneRange()
paragraph=range.OwnerParagraph

#Return the zero-based index of the specified paragraph.
body=paragraph.OwnerTextBody
index=body.ChildObjects.IndexOf(paragraph)

#Create a new table.
table=section.AddTable(True)
table.ResetCells(3,3)

#Remove the paragraph and insert table into the collection at the specified index.
body.ChildObjects.Remove(paragraph)
body.ChildObjects.Insert(index,table)

#Save to file.
document.SaveToFile(outputFile,FileFormat.Docx2013)
document.Close()