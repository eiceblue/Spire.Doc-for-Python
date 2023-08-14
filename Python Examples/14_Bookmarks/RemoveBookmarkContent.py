import unittest
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Bookmark.docx"
outputFile = "RemoveBookmarkContent.docx"


#Load the document from disk.
document = Document()
document.LoadFromFile(inputFile)

#Get the bookmark by name.            
bookmark = document.Bookmarks["Test"]

para = bookmark.BookmarkStart.Owner if isinstance(bookmark.BookmarkStart.Owner, Paragraph) else None
startIndex = para.ChildObjects.IndexOf(bookmark.BookmarkStart)
para = bookmark.BookmarkEnd.Owner if isinstance(bookmark.BookmarkEnd.Owner, Paragraph) else None
endIndex = para.ChildObjects.IndexOf(bookmark.BookmarkEnd)

#Remove the content object, and Start from next of BookmarkStart object, end up with previous of BookmarkEnd object. 
#This method is only to remove the content of the bookmark.
for i in range(startIndex + 1, endIndex):
    para.ChildObjects.RemoveAt(startIndex + 1)

#Save the document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()

