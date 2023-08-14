from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Bookmark.docx"
outputFile = "InsertImageAtBookmark.docx"

#Load the document
doc = Document()
doc.LoadFromFile(inputFile)

#Create an instance of BookmarksNavigator
bn = BookmarksNavigator(doc)

#Find a bookmark named Test
bn.MoveToBookmark("Test", True, True)

#Add a section
section0 = doc.AddSection()

#Add a paragraph for the section
paragraph = section0.AddParagraph()

#Add a picture into the paragraph
picture = paragraph.AppendPicture("./Data/Word.png")

#Add the paragraph at the position of bookmark
bn.InsertParagraph(paragraph)

#Remove the section0
doc.Sections.Remove(section0)

#Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()

