from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/HeaderAndFooter.docx"
outputFile = "RemoveFooter.docx"

# Load the document
doc = Document()
doc.LoadFromFile(inputFile)

# Get the first section
section = doc.Sections[0]

# Traverse the word document and clear all footers in different type
for i in range(section.Paragraphs.Count):
    para = section.Paragraphs.get_Item(i)
    for j in range(para.ChildObjects.Count):
        obj = para.ChildObjects.get_Item(j)
        # Clear footer in the first page
        footer = None
        footer = section.HeadersFooters[HeaderFooterType.FooterFirstPage]
        if footer is not None:
            footer.ChildObjects.Clear()
        # Clear footer in the odd page
        footer = section.HeadersFooters[HeaderFooterType.FooterOdd]
        if footer is not None:
            footer.ChildObjects.Clear()
        # Clear footer in the even page
        footer = section.HeadersFooters[HeaderFooterType.FooterEven]
        if footer is not None:
            footer.ChildObjects.Clear()

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
