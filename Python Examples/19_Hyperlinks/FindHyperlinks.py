from spire.doc import *
from spire.doc.common import *


def WriteAllText(fpath: str, content: str):
    with open(fpath, 'w') as fp:
        fp.write(content)


outputFile = "FindHyperlinks.txt"
inputFile = "./Data/Hyperlinks.docx"

# Load Document
doc = Document()
doc.LoadFromFile(inputFile)

# Create a hyperlink list
hyperlinks = []
hyperlinksText = ''
# Iterate through the items in the sections to find all hyperlinks
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
                        # Get the hyperlink text
                        hyperlinksText += field.FieldText + "\r\n"

# Save the text of all hyperlinks to TXT File and launch it
WriteAllText(outputFile, hyperlinksText)
doc.Close()
