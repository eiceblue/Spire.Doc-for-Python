
from spire.doc import *
from spire.doc.common import *

outputFile = "SetCaptionWithChapterNumber.docx"
inputFile = "./Data/SetCaptionWithChapterNumber.docx"

#Create word document
document = Document()

#Load file
document.LoadFromFile(inputFile)

#Get the first section
section = document.Sections[0]

#Label name
name = "Caption "
for i in range(section.Body.Paragraphs.Count):
    for j in range(section.Body.Paragraphs[i].ChildObjects.Count):
        if isinstance(section.Body.Paragraphs[i].ChildObjects[j], DocPicture):
            pic1 = section.Body.Paragraphs[i].ChildObjects[j] if isinstance(section.Body.Paragraphs[i].ChildObjects[j], DocPicture) else None
            body = pic1.OwnerParagraph.Owner if isinstance(pic1.OwnerParagraph.Owner, Body) else None
            if body is not None:
                imageIndex = body.ChildObjects.IndexOf(pic1.OwnerParagraph)
                #Create a new paragraph
                para = Paragraph(document)
                #Set label
                para.AppendText(name)
                #Add caption
                field1 = para.AppendField("test", FieldType.FieldStyleRef)
                #Chapter number
                field1.Code = " STYLEREF 1 \\s "
                #Chapter delimiter
                para.AppendText(" - ")
                #Add picture sequence number
                field2 = para.AppendField(name, FieldType.FieldSequence)
                field2.CaptionName = name
                field2.NumberFormat = CaptionNumberingFormat.Number
                body.Paragraphs.Insert(imageIndex + 1, para)

#Set update fields
document.IsUpdateFields = True

#Save the result file
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()