from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Text1.docx"
outputFile = "ChangeCase.docx"
     
# Create a new document and load from file
doc = Document()
doc.LoadFromFile(inputFile)
textRange = None
#Get the first paragraph and set its CharacterFormat to AllCaps
para1 = doc.Sections[0].Paragraphs[1]

for i in range(para1.ChildObjects.Count):
    obj = para1.ChildObjects.get_Item(i)
    if isinstance(obj, TextRange):
        textRange = obj if isinstance(obj, TextRange) else None
        textRange.CharacterFormat.AllCaps = True

#Get the third paragraph and set its CharacterFormat to IsSmallCaps
para2 = doc.Sections[0].Paragraphs[3]
for i in range(para2.ChildObjects.Count):
    obj = para2.ChildObjects.get_Item(i)
    if isinstance(obj, TextRange):
        textRange = obj if isinstance(obj, TextRange) else None
        textRange.CharacterFormat.IsSmallCaps = True

#Save and launch the document
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()
