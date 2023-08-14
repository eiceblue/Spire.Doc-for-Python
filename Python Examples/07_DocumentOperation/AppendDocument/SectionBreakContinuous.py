from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Sample_two sections.docx"
outputFile = "SectionBreakContinuous.docx"

#Open a Word document
doc = Document()
doc.LoadFromFile(inputFile)
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    #Set section break as continuous
    section.BreakCode = SectionBreakType.NoBreak
#Save the file
doc.SaveToFile(outputFile)
doc.Close()