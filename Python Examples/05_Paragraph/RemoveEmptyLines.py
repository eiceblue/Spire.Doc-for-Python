from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_3.docx"
outputFile = "RemoveEmptyLines.docx"

#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)

#Traverse every section on the word document and remove the null and empty paragraphs.
for k in range(document.Sections.Count):
    section = document.Sections.get_Item(k)
    i = 0
    while i < section.Body.ChildObjects.Count:
        if section.Body.ChildObjects[i].DocumentObjectType == DocumentObjectType.Paragraph:
            objItem = section.Body.ChildObjects[i]
            if isinstance(objItem, Paragraph):
                paraObj = Paragraph(objItem)
                if len(paraObj.Text) == 0:
                    section.Body.ChildObjects.Remove(objItem)
                    i -= 1

        i += 1

#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()