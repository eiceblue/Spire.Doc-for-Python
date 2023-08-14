from spire.doc import *
from spire.doc.common import *


inputFile1 = "./Data/Template.docx"
inputFile2 = "./Data/E-iceblue.png"
outputFile = "ReplaceWithImage.docx"
       
#Load Document
doc = Document()
doc.LoadFromFile(inputFile1)
#Find the string "E-iceblue" in the document
selections = doc.FindAllString("E-iceblue", True, True)
index = 0
testRange = None

#Remove the text and replace it with Image
for selection in selections:
    pic = DocPicture(doc)
    pic.LoadImage(inputFile2)

    testRange = selection.GetAsOneRange()
    index = testRange.OwnerParagraph.ChildObjects.IndexOf(testRange)
    testRange.OwnerParagraph.ChildObjects.Insert(index, pic)
    testRange.OwnerParagraph.ChildObjects.Remove(testRange)

#Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
