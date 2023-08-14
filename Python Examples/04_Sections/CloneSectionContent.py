from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/SectionTemplate.docx"
outputFile = "CloneSectionContent.docx"
   

#Load the Word document from disk
doc = Document()
doc.LoadFromFile(inputFile)

#Get the first section
sec1 = doc.Sections[0]
#Get the second section
sec2 = doc.Sections[1]

#Loop through the contents of sec1
for i in range(sec1.Body.ChildObjects.Count):
    obj = sec1.Body.ChildObjects.get_Item(i)
    #Clone the contents to sec2
    sec2.Body.ChildObjects.Add(obj.Clone())

#Save the Word document
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()