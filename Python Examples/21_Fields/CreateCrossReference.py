from spire.doc import *
from spire.doc.common import *

outputFile = "CreateCrossReference.docx"

# Create Word document.
document = Document()

# Add a new section.
section = document.AddSection()

# Create a bookmark.
paragraph = section.AddParagraph()
paragraph.AppendBookmarkStart("MyBookmark")
paragraph.AppendText("Text inside a bookmark")
paragraph.AppendBookmarkEnd("MyBookmark")

# Insert line breaks.
for i in range(0, 4):
    paragraph.AppendBreak(BreakType.LineBreak)

# Create a cross-reference field, and link it to bookmark.
field = Field(document)
field.Type = FieldType.FieldRef
field.Code = """REF MyBookmark \\p \\h"""

# Insert field to paragraph.
paragraph = section.AddParagraph()
paragraph.AppendText("For more information, see ")
paragraph.ChildObjects.Add(field)

# Insert FieldSeparator object.
fieldSeparator = FieldMark(document, FieldMarkType.FieldSeparator)
paragraph.ChildObjects.Add(fieldSeparator)

# Set display text of the field.
tr = TextRange(document)
tr.Text = "above"
paragraph.ChildObjects.Add(tr)

# Insert FieldEnd object to mark the end of the field.
fieldEnd = FieldMark(document, FieldMarkType.FieldEnd)
paragraph.ChildObjects.Add(fieldEnd)

# Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
