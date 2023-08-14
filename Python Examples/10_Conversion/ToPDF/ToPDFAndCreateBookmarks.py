
from spire.doc import *
from spire.doc.common import *

inputFile =  "./Data/BookmarkTemplate.docx"
outputFile =  "ToPDFAndCreateBookmarks.pdf"

document = Document()
#Load the document from disk
document.LoadFromFile(inputFile)
parames = ToPdfParameterList()
#Set CreateWordBookmarks to true
parames.CreateWordBookmarks = True
#//Create bookmarks using Headings
#parames.CreateWordBookmarksUsingHeadings = true
#Create bookmarks using word bookmarks
parames.CreateWordBookmarksUsingHeadings = False
document.SaveToFile(outputFile, FileFormat.PDF)
document.Close()
