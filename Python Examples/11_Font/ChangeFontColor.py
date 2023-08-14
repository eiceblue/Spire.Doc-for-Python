import unittest
from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Sample.docx"
outputFile = "ChangeFontColor.docx"
       
#Load the document
doc = Document()
doc.LoadFromFile(inputFile)

#Get the first section and first paragraph
section = doc.Sections[0]
p1 = section.Paragraphs[0]

#Iterate through the childObjects of the paragraph 1 
for i in range(p1.ChildObjects.Count):
    childObj = p1.ChildObjects.get_Item(i)
    if isinstance(childObj, TextRange):
        #Change text color
        tr = childObj if isinstance(childObj, TextRange) else None
        tr.CharacterFormat.TextColor = Color.get_RosyBrown()

#Get the second paragraph
p2 = section.Paragraphs[1]

#Iterate through the childObjects of the paragraph 2
for i in range(p2.ChildObjects.Count):
    childObj = p2.ChildObjects.get_Item(i)
    if isinstance(childObj, TextRange):
        #Change text color
        tr = childObj if isinstance(childObj, TextRange) else None
        tr.CharacterFormat.TextColor = Color.get_DarkGreen()

#Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()

