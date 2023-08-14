from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/SampleB_2.docx"
outputFile = "CreateNestedField.docx"

# Create Word document.
document = Document()
document.LoadFromFile(inputFile)

# Get the first section
section = document.Sections[0]

paragraph = section.AddParagraph()

# Create an IF field
ifField = IfField(document)
ifField.Type = FieldType.FieldIf
ifField.Code = "IF "
paragraph.Items.Add(ifField)

# Create the embedded IF field
ifField2 = IfField(document)
ifField2.Type = FieldType.FieldIf
ifField2.Code = "IF "
paragraph.ChildObjects.Add(ifField2)
paragraph.Items.Add(ifField2)
paragraph.AppendText("\"200\" < \"50\"   \"200\" \"50\" ")
embeddedEnd = document.CreateParagraphItem(ParagraphItemType.FieldMark)
(embeddedEnd if isinstance(embeddedEnd, FieldMark)
 else None).Type = FieldMarkType.FieldEnd
paragraph.Items.Add(embeddedEnd)
ifField2.End = embeddedEnd if isinstance(embeddedEnd, FieldMark) else None

paragraph.AppendText(" > ")
paragraph.AppendText("\"100\" ")
paragraph.AppendText("\"Thanks\" ")
paragraph.AppendText("\"The minimum order is 100 units\"")
end = document.CreateParagraphItem(ParagraphItemType.FieldMark)
(end if isinstance(end, FieldMark) else None).Type = FieldMarkType.FieldEnd
paragraph.Items.Add(end)
ifField.End = end if isinstance(end, FieldMark) else None

# Update all fields in the document.
document.IsUpdateFields = True

# Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
