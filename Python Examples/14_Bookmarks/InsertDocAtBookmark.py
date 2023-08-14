from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Bookmark.docx"
outputFile = "InsertDocAtBookmark.docx"

#Create the first document
document1 = Document()

#Load the first document from disk.
document1.LoadFromFile(inputFile)

#Create the second document
document2 = Document()

#Load the second document from disk.
document2.LoadFromFile("./Data/Insert.docx")

#Get the first section of the first document 
section1 = document1.Sections[0]

#Locate the bookmark
bn = BookmarksNavigator(document1)

#Find bookmark by name
bn.MoveToBookmark("Test", True, True)

#Get bookmarkStart
start = bn.CurrentBookmark.BookmarkStart

#Get the owner paragraph
para = start.OwnerParagraph

#Get the para index
index = section1.Body.ChildObjects.IndexOf(para)

#Insert the paragraphs of document2
for i in range(document2.Sections.Count):
    section2 = document2.Sections.get_Item(i)
    for j in range(section2.Paragraphs.Count):
        paragraph = section2.Paragraphs.get_Item(j)
        cloneP = paragraph.Clone()
        section1.Body.ChildObjects.Insert(index + 1, cloneP if isinstance(cloneP, Paragraph) else None)

#Save the document.
document1.SaveToFile(outputFile, FileFormat.Docx)
document1.Close()
