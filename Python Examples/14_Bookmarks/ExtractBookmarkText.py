from io import FileIO
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/BookmarkTemplate.docx"
outputFile = "ExtractBookmarkText.txt"

#Load Document
doc =Document()
doc.LoadFromFile(inputFile)

#Creates a BookmarkNavigator instance to access the bookmark
navigator = BookmarksNavigator(doc)
#Locate a specific bookmark by bookmark name
navigator.MoveToBookmark("Content")
textBodyPart = navigator.GetBookmarkContent()

#Iterate through the items in the bookmark content to get the text
text = ''
for i in range(textBodyPart.BodyItems.Count):
    item = textBodyPart.BodyItems.get_Item(i)
    if isinstance(item, Paragraph):
        for j in range(( item if isinstance(item, Paragraph) else None).ChildObjects.Count):
            childObject = ( item if isinstance(item, Paragraph) else None).ChildObjects.get_Item(j)
            if isinstance(childObject, TextRange):
                text += ( childObject if isinstance(childObject, TextRange) else None).Text

#Write the contents in a TXT file
with FileIO(outputFile, mode="w") as f:
    f.write(text.encode("utf-8"))
doc.Close()

