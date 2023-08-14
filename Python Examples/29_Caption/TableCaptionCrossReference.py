from spire.doc import *
from spire.doc.common import *

outputFile = "TableCaptionCrossReference.docx"

#Create word document
document = Document()

#Get the first section
section = document.AddSection()

#Create a table
table = section.AddTable(True)
table.ResetCells(2, 3)

#Add caption to the table
captionParagraph = table.AddCaption("Table", CaptionNumberingFormat.Number, CaptionPosition.BelowItem)

#Create a bookmark
bookmarkName = "Table_1"
paragraph = section.AddParagraph()
paragraph.AppendBookmarkStart(bookmarkName)
paragraph.AppendBookmarkEnd(bookmarkName)

#Replace bookmark content
navigator = BookmarksNavigator(document)
navigator.MoveToBookmark(bookmarkName)
part = navigator.GetBookmarkContent()
part.BodyItems.Clear()
part.BodyItems.Add(captionParagraph)
navigator.ReplaceBookmarkContent(part)

#Create cross-reference field to point to bookmark "Table_1"
field = Field(document)
field.Type = FieldType.FieldRef
field.Code = """REF Table_1 \\p \\h"""

#Insert line breaks
for i in range(0, 3):
    paragraph.AppendBreak(BreakType.LineBreak)

#Insert field to paragraph
paragraph = section.AddParagraph()
testRange = paragraph.AppendText("This is a table caption cross-reference, ")
testRange.CharacterFormat.FontSize = 14
paragraph.ChildObjects.Add(field)

#Insert FieldSeparator object
fieldSeparator = FieldMark(document, FieldMarkType.FieldSeparator)
paragraph.ChildObjects.Add(fieldSeparator)

#Set display text of the field
tr = TextRange(document)
tr.Text = "Table 1"
tr.CharacterFormat.FontSize = 14
tr.CharacterFormat.TextColor = Color.get_DeepSkyBlue()
paragraph.ChildObjects.Add(tr)

#Insert FieldEnd object to mark the end of the field
fieldEnd = FieldMark(document, FieldMarkType.FieldEnd)
paragraph.ChildObjects.Add(fieldEnd)

#Update fields
document.IsUpdateFields = True

#Save the Word document
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()