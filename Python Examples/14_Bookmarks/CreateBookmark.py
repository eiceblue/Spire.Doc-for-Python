from spire.doc import *
from spire.doc.common import *

outputFile = "CreateBookmark.docx"
def _CreateBookmark(section):
    paragraph = section.AddParagraph()
    txtRange = paragraph.AppendText("The following example demonstrates how to create bookmark in a Word document.")
    txtRange.CharacterFormat.Italic = True

    section.AddParagraph()
    paragraph = section.AddParagraph()
    txtRange = paragraph.AppendText("Simple Create Bookmark.")
    txtRange.CharacterFormat.TextColor = Color.get_CornflowerBlue()
    paragraph.ApplyStyle(BuiltinStyle.Heading2)

    #Write simple CreateBookmarks.
    section.AddParagraph()
    paragraph = section.AddParagraph()
    paragraph.AppendBookmarkStart("SimpleCreateBookmark")
    paragraph.AppendText("This is a simple bookmark.")
    paragraph.AppendBookmarkEnd("SimpleCreateBookmark")

    section.AddParagraph()
    paragraph = section.AddParagraph()
    txtRange = paragraph.AppendText("Nested Create Bookmark.")
    txtRange.CharacterFormat.TextColor = Color.get_CornflowerBlue()
    paragraph.ApplyStyle(BuiltinStyle.Heading2)

    #Write nested CreateBookmarks.
    section.AddParagraph()
    paragraph = section.AddParagraph()
    paragraph.AppendBookmarkStart("Root")
    txtRange = paragraph.AppendText(" This is Root data ")
    txtRange.CharacterFormat.Italic = True
    paragraph.AppendBookmarkStart("NestedLevel1")
    txtRange = paragraph.AppendText(" This is Nested Level1 ")
    txtRange.CharacterFormat.Italic = True
    txtRange.CharacterFormat.TextColor = Color.get_DarkSlateGray()
    paragraph.AppendBookmarkStart("NestedLevel2")
    txtRange = paragraph.AppendText(" This is Nested Level2 ")
    txtRange.CharacterFormat.Italic = True
    txtRange.CharacterFormat.TextColor = Color.get_DimGray()
    paragraph.AppendBookmarkEnd("NestedLevel2")
    paragraph.AppendBookmarkEnd("NestedLevel1")
    paragraph.AppendBookmarkEnd("Root")

#Create word document.
document = Document()

#Create a new section.
section = document.AddSection()

_CreateBookmark(section)

#Save the document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()




