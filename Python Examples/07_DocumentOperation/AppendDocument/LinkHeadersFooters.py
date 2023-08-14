from spire.doc import *
from spire.doc.common import *

inputFile1 = "./Data/Template_N1.docx"
inputFile2 = "./Data/Template_N2.docx"
outputFile = "LinkHeadersFooters.docx"

#Load the source file
srcDoc = Document()
srcDoc.LoadFromFile(inputFile1)
#Load the destination file
dstDoc = Document()
dstDoc.LoadFromFile(inputFile2)
#Link the headers and footers in the source file
srcDoc.Sections[0].HeadersFooters.Header.LinkToPrevious = True
srcDoc.Sections[0].HeadersFooters.Footer.LinkToPrevious = True
#Clone the sections of source to destination
for i in range(srcDoc.Sections.Count):
    section = srcDoc.Sections.get_Item(i)
    dstDoc.Sections.Add(section.Clone())  
#Save the document
dstDoc.SaveToFile(outputFile, FileFormat.Docx2013)
srcDoc.Close()
dstDoc.Close()