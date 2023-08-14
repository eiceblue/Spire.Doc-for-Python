from spire.doc import *
from spire.doc.common import *

class StructureTags:
    def __init__(self):
        #instance fields found by C# to Python Converter:
        self._m_tagInlines = None
        self._m_tags = None

    def get_tag_inlines(self):
        if self._m_tagInlines is None:
            self._m_tagInlines = []
        return self._m_tagInlines
    def set_tag_inlines(self, value):
        self._m_tagInlines = value
    def get_tags(self):
        if self._m_tags is None:
            self._m_tags = []
        return self._m_tags
    def set_tags(self, value):
        self._m_tags = value

def _GetAllTags(document):
    structureTags = StructureTags()
    for i in range(document.Sections.Count):
        section = document.Sections.get_Item(i)
        for j in range(section.Body.ChildObjects.Count):
            obj = section.Body.ChildObjects.get_Item(j)
            if obj.DocumentObjectType == DocumentObjectType.StructureDocumentTag:
                structureTags.get_tags().append( obj if isinstance(obj, StructureDocumentTag) else None)

            elif obj.DocumentObjectType == DocumentObjectType.Paragraph:
                tempPara = obj if isinstance(obj, Paragraph) else None
                for k in range(tempPara.ChildObjects.Count):
                    pobj = tempPara.ChildObjects.get_Item(k)
                    if pobj.DocumentObjectType == DocumentObjectType.StructureDocumentTagInline:
                        tempPobj = pobj if isinstance(pobj, StructureDocumentTagInline) else None
                        structureTags.get_tag_inlines().append(tempPobj)
            elif obj.DocumentObjectType == DocumentObjectType.Table:
                tempTable = obj if isinstance(obj, Table) else None
                for x in range(tempTable.Rows.Count):
                    row = tempTable.Rows.get_Item(x)
                    for g in range(row.Cells.Count):
                        cell = row.Cells.get_Item(g)
                        for z in range(cell.ChildObjects.Count):
                            cellChild = cell.ChildObjects.get_Item(z)
                            if cellChild.DocumentObjectType == DocumentObjectType.StructureDocumentTag:
                                structureTags.get_tags().append( cellChild if isinstance(cellChild, StructureDocumentTag) else None)
                            elif cellChild.DocumentObjectType == DocumentObjectType.Paragraph:
                                tempParagraph = cellChild if isinstance(cellChild, Paragraph) else None
                            for p in range(tempParagraph.ChildObjects.Count):
                                pobj = tempParagraph.ChildObjects.get_Item(p)
                                if pobj.DocumentObjectType == DocumentObjectType.StructureDocumentTagInline:
                                    structureTags.get_tag_inlines().append( pobj if isinstance(pobj, StructureDocumentTagInline) else None)
    return structureTags


def WriteAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s)
        

inputFile = "./Data/ContentControl.docx"
outputFile = "GetContentControlProperty.txt"

#Create a new document and load from file
doc = Document()
doc.LoadFromFile(inputFile)

#Get all structureTags in the Word document
structureTags = _GetAllTags(doc)
#Get all StructureDocumentTagInline objects
tagInlines = structureTags.get_tag_inlines()
strProperty = ''
strProperty += "Alias of contentControl" + "\t" + "ID  " + "\t" + "Tag     " + "\t" + "STDType" + "\r\n"
#Get properties of all tagInlines
for i, unusedItem in enumerate(tagInlines):
    alias = tagInlines[i].SDTProperties.Alias
    objId = tagInlines[i].SDTProperties.Id
    tag = tagInlines[i].SDTProperties.Tag
    STDType = str(tagInlines[i].SDTProperties.SDTType)
    strProperty += alias + ",\t" + str(objId) + ",\t" + tag + ",\t" + STDType + "\r\n"

#Get all StructureDocumentTag objects
tags = structureTags.get_tags()
#Get properties of all tags
for i, unusedItem in enumerate(tags):
    alias = tags[i].SDTProperties.Alias
    objId = tags[i].SDTProperties.Id
    tag = tags[i].SDTProperties.Tag
    STDType = str(tags[i].SDTProperties.SDTType)
    strProperty += alias + ",\t" + str(objId) + ",\t" + tag + ",\t" + STDType + "\r\n"

#Save the property to a text document and launch it
WriteAllText(outputFile, strProperty)
doc.Close()



        