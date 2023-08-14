from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_4.docx"
outputFolder = "SplitDocBySectionBreak/"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Define another new word document object.
newWord = None
#Split a Word document into multiple documents by section break.
for i in range(document.Sections.Count):
    unusedItem = document.Sections.get_Item(i)
    result = outputFolder+"SplitDocBySectionBreak_{0}.docx".format(i)
    newWord = Document()
    newWord.Sections.Add(document.Sections[i].Clone())
    #Save the file.
    newWord.SaveToFile(result)
    newWord.Close()