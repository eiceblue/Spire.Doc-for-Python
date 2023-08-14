from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Bookmark.docx"
outputFile = "FromBookmark.docx"

#Create the source document
sourcedocument = Document()

#Load the source document from disk.
sourcedocument.LoadFromFile(inputFile)

#Create a destination document
destinationDoc = Document()

#Add a section for destination document
section = destinationDoc.AddSection()

#Add a paragraph for destination document
paragraph = section.AddParagraph()

#Locate the bookmark in source document
navigator = BookmarksNavigator(sourcedocument)

#Find bookmark by name
navigator.MoveToBookmark("Test", True, True)

#get text body part
textBodyPart = navigator.GetBookmarkContent()

#Create a TextRange type list
listTextRange = []

#Traverse the items of text body
for i in range(textBodyPart.BodyItems.Count):
    item = textBodyPart.BodyItems.get_Item(i)
    #if it is paragraph
    if isinstance(item, Paragraph):
        tempItems = (item if isinstance(item, Paragraph) else None).ChildObjects
#Traverse the ChildObjects of the paragraph
for j in range(tempItems.Count):
    childObject = tempItems.get_Item(j)
    #if it is TextRange
    if isinstance(childObject, TextRange):
        #Add it into list
        textRange = childObject if isinstance(childObject, TextRange) else None
        listTextRange.append(textRange)

#Add the extract content to destinationDoc document
for m, unusedItem in enumerate(listTextRange):
    paragraph.Items.Add(listTextRange[m].Clone())

#Save the document.
destinationDoc.SaveToFile(outputFile, FileFormat.Docx)
destinationDoc.Close()
