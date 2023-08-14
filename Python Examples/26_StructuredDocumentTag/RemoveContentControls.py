from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/RemoveContentControls.docx"
outputFile = "RemoveContentControls.docx"

#Load document from disk
doc = Document()
doc.LoadFromFile(inputFile)

#Loop through sections
for s in range(doc.Sections.Count):
    section = doc.Sections[s]
    i = 0
    while i < section.Body.ChildObjects.Count:
        #Loop through contents in paragraph
        if isinstance(section.Body.ChildObjects[i], Paragraph):
            para = section.Body.ChildObjects[i] if isinstance(section.Body.ChildObjects[i], Paragraph) else None
            j = 0
            while j < para.ChildObjects.Count:
                #Find the StructureDocumentTagInline
                if isinstance(para.ChildObjects[j], StructureDocumentTagInline):
                    sdt = para.ChildObjects[j] if isinstance(para.ChildObjects[j], StructureDocumentTagInline) else None
                    #Remove the content control from paragraph
                    para.ChildObjects.Remove(sdt)
                    j -= 1
                j += 1
        if isinstance(section.Body.ChildObjects[i], StructureDocumentTag):
            sdt = section.Body.ChildObjects[i] if isinstance(section.Body.ChildObjects[i], StructureDocumentTag) else None
            section.Body.ChildObjects.Remove(sdt)
            i -= 1
        i += 1

#Save the Word document
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()
