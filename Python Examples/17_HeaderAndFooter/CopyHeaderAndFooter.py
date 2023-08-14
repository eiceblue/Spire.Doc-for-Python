from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/HeaderAndFooter.docx"
inputFile_1 = "./Data/Template.docx"
outputFile = "CopyHeaderAndFooter.docx"

# Load the source file
doc1 = Document()
doc1.LoadFromFile(inputFile)

# Get the header section from the source document
header = doc1.Sections[0].HeadersFooters.Header

# Load the destination file
doc2 = Document()
doc2.LoadFromFile(inputFile_1)

# Copy each object in the header of source file to destination file
for i in range(doc2.Sections.Count):
    section = doc2.Sections.get_Item(i)
    for j in range(header.ChildObjects.Count):
        obj = header.ChildObjects.get_Item(j)
        section.HeadersFooters.Header.ChildObjects.Add(obj.Clone())

# Save and launch document
doc2.SaveToFile(outputFile, FileFormat.Docx)
doc1.Close()
doc2.Close()
