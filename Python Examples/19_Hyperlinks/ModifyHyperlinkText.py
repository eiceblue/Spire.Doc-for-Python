from spire.doc import *
from spire.doc.common import *

outputFile = "ModifyHyperlinkText.docx"
inputFile = "./Data/Hyperlinks.docx"

# Load Document
doc = Document()
doc.LoadFromFile(inputFile)

# Find all hyperlinks in the Word document
hyperlinks = []
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    for j in range(section.Body.ChildObjects.Count):
        sec = section.Body.ChildObjects.get_Item(j)
        if sec.DocumentObjectType == DocumentObjectType.Paragraph:
            for k in range((sec if isinstance(sec, Paragraph) else None).ChildObjects.Count):
                para = (sec if isinstance(sec, Paragraph)
                        else None).ChildObjects.get_Item(k)
                if para.DocumentObjectType == DocumentObjectType.Field:
                    field = para if isinstance(para, Field) else None

                    if field.Type == FieldType.FieldHyperlink:
                        hyperlinks.append(field)

# Reset the property of hyperlinks[0].FieldText by using the index of the hyperlink
hyperlinks[0].FieldText = "Spire.Doc component"

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
