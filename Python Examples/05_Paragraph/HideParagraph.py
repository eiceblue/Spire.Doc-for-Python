from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Template_Docx_1.docx"
outputFile = "HideParagraph.docx"

#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)

#Get the first section and the first paragraph from the word document.
sec = document.Sections[0]
para = sec.Paragraphs[0]

#Loop through the textranges and set CharacterFormat.Hidden property as true to hide the texts.
for i in range(para.ChildObjects.Count):
    obj = para.ChildObjects.get_Item(i)
    if isinstance(obj, TextRange):
        trange = obj if isinstance(obj, TextRange) else None
        trange.CharacterFormat.Hidden = True

#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()