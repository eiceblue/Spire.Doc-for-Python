from spire.doc import *
from spire.doc.common import *

def WriteAllText(fname:str,text:List[str]):
        fp = open(fname,"w")
        for s in text:
            fp.write(s)
        fp.close()

inputFile = "./Data/Sample.docx"
outputFile = "RecurseAllDocumentObject.txt"

builder = ""
#Create Word document.
document = Document()
document.LoadFromFile(inputFile)
#find all document object
for i in range(document.Sections.Count):
    section = document.Sections.get_Item(i)
    SectionIndex = document.GetIndex(section)
    builder += "section index {} has following ChildObjects".format(SectionIndex)
    builder += "\n"
    for j in range(section.Body.ChildObjects.Count):
        obj = section.Body.ChildObjects.get_Item(j)
        builder += "Index : {}, ChildObject Type: {}".format(section.Body.GetIndex(obj), obj.DocumentObjectType.name)
        builder += "\n"
        
        if obj.DocumentObjectType is DocumentObjectType.Paragraph:
            paragraph = obj if isinstance(obj, Paragraph) else None
            builder += "\tParagraph index {} has following ChildObjects".format(section.Body.GetIndex(paragraph))
            builder += "\n"
            for k in range(paragraph.ChildObjects.Count):
                obj2 = paragraph.ChildObjects.get_Item(k)
                builder += "\tIndex : {}, ChildObject Type: {}".format(paragraph.GetIndex(obj2), obj2.DocumentObjectType.name)
                builder += "\n"
    builder += " "
    builder += "\n"

WriteAllText(outputFile, builder)