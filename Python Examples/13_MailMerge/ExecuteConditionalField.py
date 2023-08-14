from spire.doc import *
from spire.doc.common import *


def _CreateIFField1(document, paragraph):
    ifField = IfField(document)
    ifField.Type = FieldType.FieldIf
    ifField.Code = "IF "
    paragraph.Items.Add(ifField)

    paragraph.AppendField("Count", FieldType.FieldMergeField)
    paragraph.AppendText(" > ")
    paragraph.AppendText("\"1\" ")
    paragraph.AppendText("\"Greater than one\" ")
    paragraph.AppendText("\"Less than one\"")

    end = document.CreateParagraphItem(ParagraphItemType.FieldMark)
    tempFieldMark = ( end if isinstance(end, FieldMark) else None)
    if tempFieldMark != None:
        tempFieldMark.Type = FieldMarkType.FieldEnd
        
    paragraph.Items.Add(end)
    ifField.End = end if isinstance(end, FieldMark) else None

def _CreateIFField2(document, paragraph):
    ifField = IfField(document)
    ifField.Type = FieldType.FieldIf
    ifField.Code = "IF "
    paragraph.Items.Add(ifField)

    paragraph.AppendField("Age", FieldType.FieldMergeField)
    paragraph.AppendText(" > ")
    paragraph.AppendText("\"50\" ")
    paragraph.AppendText("\"The old man\" ")
    paragraph.AppendText("\"The young man\"")

    end = document.CreateParagraphItem(ParagraphItemType.FieldMark)
    tempFieldMark = ( end if isinstance(end, FieldMark) else None)
    tempFieldMark.Type = FieldMarkType.FieldEnd
    paragraph.Items.Add(end)

    ifField.End = end if isinstance(end, FieldMark) else None


outputFile = "ExecuteConditionalField.docx"

doc = Document()
#Add a new section 
section = doc.AddSection()
#Add a new paragraph for a section 
paragraph = section.AddParagraph()

_CreateIFField1(doc, paragraph)
paragraph = section.AddParagraph()
_CreateIFField2(doc, paragraph)

fieldName = ["Count", "Age"]
fieldValue = ["2", "30"]

doc.MailMerge.Execute(fieldName, fieldValue)
doc.IsUpdateFields = True

doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()


