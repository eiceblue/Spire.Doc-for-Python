from spire.doc import *
from spire.doc.common import *


class StructureTags:

    def __init__(self):
        self._m_tagInlines = None

    def get_tag_inlines(self):
        if self._m_tagInlines is None:
            self._m_tagInlines = []
        return self._m_tagInlines
    def set_tag_inlines(self, value):
        self._m_tagInlines = value

def _GetAllTags(document):

    #Create StructureTags
    structureTags = StructureTags()

    #Travel document sections
    for i in range(document.Sections.Count):
        section = document.Sections.get_Item(i)
        for j in range(section.Body.ChildObjects.Count):
            obj = section.Body.ChildObjects.get_Item(j)
            #Travel document paragraphs
            if obj.DocumentObjectType == DocumentObjectType.Paragraph:
                tempParagraph = ( obj if isinstance(obj, Paragraph) else None)
                for k in range(tempParagraph.ChildObjects.Count):
                    pobj = tempParagraph.ChildObjects.get_Item(k)
                    #Get StructureDocumentTagInline
                    if pobj.DocumentObjectType == DocumentObjectType.StructureDocumentTagInline:
                        structureTags.get_tag_inlines().append( pobj if isinstance(pobj, StructureDocumentTagInline) else None)

    return structureTags



inputFile = "./Data/CheckBoxContentControl.docx"
outputFile = "UpdateCheckBox.docx"
      
#Create a document
document = Document()

#Load the document from disk.
document.LoadFromFile(inputFile)

#Call StructureTags
structureTags = _GetAllTags(document)

#Create list 
tagInlines = structureTags.get_tag_inlines()

#Get the controls
for i, item in enumerate(tagInlines):
    #Get the type
    sdtType = item.SDTProperties.SDTType.name

    #Update the status
    if sdtType == "CheckBox":
        tempPro = item.SDTProperties.ControlProperties
        scb = tempPro if isinstance(tempPro, SdtCheckBox) else None
        if scb.Checked:
            scb.Checked = False
        else:
            scb.Checked = True

#Save the document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()

