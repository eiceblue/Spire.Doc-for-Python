from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Bookmark.docx"
outputFile = "CopyBookmarkContent.docx"

#Load the document from disk.
doc = Document()
doc.LoadFromFile(inputFile)

#Get the bookmark by name.
bookmark = doc.Bookmarks["Test"]
docObj = None

#Judge if the paragraph includes the bookmark exists in the table, if it exists in cell,
#Then need to find its outermost parent object(Table),
#and get the start/end index of current object on body.
if ( bookmark.BookmarkStart.Owner if isinstance(bookmark.BookmarkStart.Owner, Paragraph) else None).IsInCell:
    docObj = bookmark.BookmarkStart.Owner.Owner.Owner.Owner
else:
    docObj = bookmark.BookmarkStart.Owner
startIndex = doc.Sections[0].Body.ChildObjects.IndexOf(docObj)
if ( bookmark.BookmarkEnd.Owner if isinstance(bookmark.BookmarkEnd.Owner, Paragraph) else None).IsInCell:
    docObj = bookmark.BookmarkEnd.Owner.Owner.Owner.Owner
else:
    docObj = bookmark.BookmarkEnd.Owner
endIndex = doc.Sections[0].Body.ChildObjects.IndexOf(docObj)

#Get the start/end index of the bookmark object on the paragraph.
para = bookmark.BookmarkStart.Owner if isinstance(bookmark.BookmarkStart.Owner, Paragraph) else None
pStartIndex = para.ChildObjects.IndexOf(bookmark.BookmarkStart)
para = bookmark.BookmarkEnd.Owner if isinstance(bookmark.BookmarkEnd.Owner, Paragraph) else None
pEndIndex = para.ChildObjects.IndexOf(bookmark.BookmarkEnd)

#Get the content of current bookmark and copy.
select = TextBodySelection(doc.Sections[0].Body, startIndex, endIndex, pStartIndex, pEndIndex)
body = TextBodyPart(select)
for i in range(body.BodyItems.Count):
    doc.Sections[0].Body.ChildObjects.Add(body.BodyItems[i].Clone())


#Save the document.
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()