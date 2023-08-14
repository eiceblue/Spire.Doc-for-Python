from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Bookmark.docx"
outputFile = "ReplaceBookmarkContent.docx"
       
#Load the document from disk.
doc = Document()
doc.LoadFromFile(inputFile)

#Locate the bookmark.
bookmarkNavigator = BookmarksNavigator(doc)
bookmarkNavigator.MoveToBookmark("Test")

#Replace the context with new.
bookmarkNavigator.ReplaceBookmarkContent("This is replaced content.", False)

#Save the document.
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()

