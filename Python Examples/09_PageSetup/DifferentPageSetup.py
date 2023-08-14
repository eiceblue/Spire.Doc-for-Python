from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/DifferentPageSetup.docx"
outputFile = "DifferentPageSetup.docx"

#Open a Word document
doc = Document()
doc.LoadFromFile(inputFile)

#Get the second section 
SectionTwo = doc.Sections[1]

#Set the orientation
SectionTwo.PageSetup.Orientation = PageOrientation.Landscape

#Set page size
#SectionTwo.PageSetup.PageSize = new SizeF(800, 800)

doc.SaveToFile(outputFile)
doc.Close()
