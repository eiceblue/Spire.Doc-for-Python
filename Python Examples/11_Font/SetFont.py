import unittest
from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Sample.docx"
outputFile = "SetFont.docx"  
#Load the document
doc = Document()
doc.LoadFromFile(inputFile)

#Get the first section 
s = doc.Sections[0]

#Get the second paragraph
p = s.Paragraphs[1]

#Create a characterFormat object
characterFormat = CharacterFormat(doc)
#Set font
characterFormat.FontName = "Arial"
characterFormat.FontSize = 16

#Loop through the childObjects of paragraph 
for i in range(p.ChildObjects.Count):
    childObj = p.ChildObjects.get_Item(i)
    if isinstance(childObj, TextRange):
        #Apply character format
        tr = childObj if isinstance(childObj, TextRange) else None
        tr.ApplyCharacterFormat(characterFormat)

#Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
