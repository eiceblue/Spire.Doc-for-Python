from spire.doc import *
from spire.doc.common import *


outputFile = "CreateBookmarkForTable.docx"
def _CreateBookmarkForTable(doc, section):
    #Add a paragraph
    paragraph = section.AddParagraph()

    #Append text for added paragraph
    txtRange = paragraph.AppendText("The following example demonstrates how to create bookmark for a table in a Word document.")

    #Set the font in italic
    txtRange.CharacterFormat.Italic = True

    #Append bookmark start
    paragraph.AppendBookmarkStart("CreateBookmark")

    #Append bookmark end
    paragraph.AppendBookmarkEnd("CreateBookmark")

    #Add table
    table = section.AddTable(True)

    #Set the number of rows and columns
    table.ResetCells(2, 2)

    #Append text for table cells
    range = table.Rows.get_Item(0).Cells.get_Item(0).AddParagraph().AppendText("sampleA")
    range = table.Rows.get_Item(0).Cells.get_Item(1).AddParagraph().AppendText("sampleB")
    range = table.Rows.get_Item(1).Cells.get_Item(0).AddParagraph().AppendText("120")
    range = table.Rows.get_Item(1).Cells.get_Item(1).AddParagraph().AppendText("260")

    #Get the bookmark by index.
    bookmark = doc.Bookmarks[0]

    #Get the name of bookmark.
    bookmarkName = bookmark.Name

    #Locate the bookmark by name.
    navigator = BookmarksNavigator(doc)
    navigator.MoveToBookmark(bookmarkName)

    #Add table to TextBodyPart
    part = navigator.GetBookmarkContent()
    part.BodyItems.Add(table)

    #Replace bookmark cotent with table
    navigator.ReplaceBookmarkContent(part)


#Create word document.
document = Document()

#Add a new section.
section = document.AddSection()

#Create bookmark for a table
_CreateBookmarkForTable(document, section)

#Save the document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()


   
