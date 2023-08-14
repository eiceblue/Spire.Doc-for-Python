from spire.doc import *
from spire.doc.common import *
inputFile = "./Data/Bookmark.docx"
outputFile = "RemoveBookmark.docx"
    
#Load the document from disk.
document = Document()
document.LoadFromFile(inputFile)

#Get the bookmark by name.
bookmark = document.Bookmarks["Test"]

#Remove the bookmark, not its content.
document.Bookmarks.Remove(bookmark)

#Save the document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
