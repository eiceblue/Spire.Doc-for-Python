from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/MultiplePages.docx"
inputFile_1 = "./Data/HeaderAndFooter.docx"
outputFile = "AddHeaderOnlyFirstPage.docx"

# Load the source file
doc1 = Document()
doc1.LoadFromFile(inputFile_1)

# Get the header from the first section
header = doc1.Sections[0].HeadersFooters.Header

# Load the destination file
doc2 = Document()
doc2.LoadFromFile(inputFile)

# Get the first page header of the destination document
firstPageHeader = doc2.Sections[0].HeadersFooters.FirstPageHeader

# Specify that the current section has a different header/footer for the first page
for i in range(doc2.Sections.Count):
    section = doc2.Sections.get_Item(i)
    section.PageSetup.DifferentFirstPageHeaderFooter = True

# Removes all child objects in firstPageHeader
firstPageHeader.Paragraphs.Clear()

# Add all child objects of the header to firstPageHeader
for j in range(header.ChildObjects.Count):
    obj = header.ChildObjects.get_Item(j)
    firstPageHeader.ChildObjects.Add(obj.Clone())

# Save and launch the file
doc2.SaveToFile(outputFile, FileFormat.Docx)
doc1.Close()
doc2.Close()
