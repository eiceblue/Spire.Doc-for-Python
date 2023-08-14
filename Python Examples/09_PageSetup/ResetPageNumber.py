
from spire.doc import *
from spire.doc.common import *

outputFile = "ResetPageNumber.docx"
#Create three Word documents and load three different Word documents from disk.
document1 = Document()
document1.LoadFromFile("./Data/ResetPageNumber1.docx")
document2 = Document()
document2.LoadFromFile("./Data/ResetPageNumber2.docx")
document3 = Document()
document3.LoadFromFile("./Data/ResetPageNumber3.docx")
#Use section method to combine all documents into one word document.
for i in range(document2.Sections.Count):
    sec = document2.Sections.get_Item(i)
    document1.Sections.Add(sec.Clone())
for i in range(document3.Sections.Count):
    sec = document3.Sections.get_Item(i)
    document1.Sections.Add(sec.Clone())
#Traverse every section of document1.
for i in range(document1.Sections.Count):
    sec = document1.Sections.get_Item(i)
    #Traverse every object of the footer.
    for j in range(sec.HeadersFooters.Footer.ChildObjects.Count):
        obj = sec.HeadersFooters.Footer.ChildObjects.get_Item(j)
        if obj.DocumentObjectType == DocumentObjectType.StructureDocumentTag:
            para = obj.ChildObjects[0]
            for k in range(para.ChildObjects.Count):
                item = para.ChildObjects.get_Item(k)
                if item.DocumentObjectType == DocumentObjectType.Field:
                    #Find the item and its field type is FieldNumPages.
                    if ( item if isinstance(item, Field) else None).Type == FieldType.FieldNumPages:
                        #Change field type to FieldSectionPages.
                        ( item if isinstance(item, Field) else None).Type = FieldType.FieldSectionPages
#Restart page number of section and set the starting page number to 1.
document1.Sections[1].PageSetup.RestartPageNumbering = True
document1.Sections[1].PageSetup.PageStartingNumber = 1
document1.Sections[2].PageSetup.RestartPageNumbering = True
document1.Sections[2].PageSetup.PageStartingNumber = 1
#Save to file.
document1.SaveToFile(outputFile, FileFormat.Docx2013)
document1.Close()
