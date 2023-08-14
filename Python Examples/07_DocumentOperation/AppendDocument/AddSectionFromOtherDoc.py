from spire.doc import *
from spire.doc.common import *

inputFile1 = "./Data/Sample.docx"
inputFile2 = "./Data/SampleB_1.docx"
outputFile = "AddSectionFromOtherDoc.docx"

#Open a Word document as target document
TarDoc = Document(inputFile1)
#Open a Word document as source document
SouDoc = Document(inputFile2)
#Get the second section from source document
Ssection = SouDoc.Sections[0]
#Add the section in target document
TarDoc.Sections.Add(Ssection.Clone())
#Save the file
TarDoc.SaveToFile(outputFile, FileFormat.Docx)
SouDoc.Close()
TarDoc.Close()