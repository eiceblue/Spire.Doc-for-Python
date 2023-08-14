from spire.doc import *
from spire.doc.common import *

outputFile = "LockContentControlContent.docx"

htmlString = "<table style=\"width: 100 % \">" + "<tr><th> Number </th><th> Name </th ><th>Age</th ></tr>" + "<tr><td> 1 </td><td> Smith </td><td> 50 </td></tr>" + "<tr> <td> 2 </td><td> Jackson </td><td> 94 </td> </tr>" + "</table>"
doc = Document()
section = doc.AddSection()
paragraph = section.AddParagraph()
paragraph.AppendHTML(htmlString)

#Create StructureDocumentTag for document
sdt = StructureDocumentTag(doc)
section2 = doc.AddSection()
section2.Body.ChildObjects.Add(sdt)

#Specify the type
sdt.SDTProperties.SDTType = SdtType.RichText
for k in range(section.Body.ChildObjects.Count):
    obj = section.Body.ChildObjects.get_Item(k)
    if obj.DocumentObjectType == DocumentObjectType.Table:
        sdt.SDTContent.ChildObjects.Add(obj.Clone())

#Lock content
sdt.SDTProperties.LockSettings = LockSettingsType.ContentLocked
doc.Sections.Remove(section)

#Save the Word document
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()
