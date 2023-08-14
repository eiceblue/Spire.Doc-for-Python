from spire.doc import *
from spire.doc.common import *


def _CreateIfField(document, paragraph):
    ifField = IfField(document)
    ifField.Type = FieldType.FieldIf
    ifField.Code = "IF "

    paragraph.Items.Add(ifField)
    paragraph.AppendField("Count", FieldType.FieldMergeField)
    paragraph.AppendText(" > ")
    paragraph.AppendText("\"100\" ")
    paragraph.AppendText("\"Thanks\" ")
    paragraph.AppendText("\"The minimum order is 100 units\"")

    end = document.CreateParagraphItem(ParagraphItemType.FieldMark)
    (end if isinstance(end, FieldMark) else None).Type = FieldMarkType.FieldEnd
    paragraph.Items.Add(end)
    ifField.End = end if isinstance(end, FieldMark) else None


outputFile = "CreateIFField.docx"

# Create Word document.
document = Document()

# Add a new section.
section = document.AddSection()

# Add a new paragraph.
paragraph = section.AddParagraph()

# Define a method of creating an IF Field.
_CreateIfField(document, paragraph)

# Define merged data.
fieldName = ["Count"]
fieldValue = ["2"]

# Merge data into the IF Field.
document.MailMerge.Execute(fieldName, fieldValue)

# Update all fields in the document.
document.IsUpdateFields = True

# Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
