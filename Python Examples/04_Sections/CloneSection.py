from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/SectionTemplate.docx"
outputFile = "CloneSection.docx"
      
#Load source file
srcDoc = Document()
srcDoc.LoadFromFile(inputFile)

#Create destination file
desDoc = Document()

cloneSection = None
for i in range(srcDoc.Sections.Count):
    section = srcDoc.Sections.get_Item(i)
    #Clone section
    cloneSection = section.Clone()
    #Add the cloneSection in destination file
    desDoc.Sections.Add(cloneSection)
#Save the Word
desDoc.SaveToFile(outputFile, FileFormat.Docx2013)
desDoc.Close()
