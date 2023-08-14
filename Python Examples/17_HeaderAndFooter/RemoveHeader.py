from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/HeaderAndFooter.docx"
outputFile = "RemoveHeader.docx"

# Load the document
doc = Document()
doc.LoadFromFile(inputFile)

# Get the first section of the document
section = doc.Sections[0]

# Traverse the word document and clear all headers in different type
for i in range(section.Paragraphs.Count):
    para = section.Paragraphs.get_Item(i)
    for j in range(para.ChildObjects.Count):
        obj = para.ChildObjects.get_Item(j)
        # Clear footer in the first page
        header = None
        header = section.HeadersFooters[HeaderFooterType.HeaderFirstPage]
        if header is not None:
            header.ChildObjects.Clear()
        # Clear footer in the odd page
        header = section.HeadersFooters[HeaderFooterType.HeaderOdd]
        if header is not None:
            header.ChildObjects.Clear()
        # Clear footer in the even page
        header = section.HeadersFooters[HeaderFooterType.HeaderEven]
        if header is not None:
            header.ChildObjects.Clear()

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
