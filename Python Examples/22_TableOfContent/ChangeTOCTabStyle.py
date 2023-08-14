from spire.doc import *
from spire.doc.common import *

outputFile = "ChangeTOCTabStyle.docx"
inputFile = "./Data/Template_Toc.docx"

#Load document from disk
doc = Document()
doc.LoadFromFile(inputFile)

#Loop through sections
for k in range(doc.Sections.Count):
    section = doc.Sections.get_Item(k)
    #Loop through content of section
    for j in range(section.Body.ChildObjects.Count):
        obj = section.Body.ChildObjects.get_Item(j)
        #Find the structure document tag
        if isinstance(obj, StructureDocumentTag):
            tag = obj if isinstance(obj, StructureDocumentTag) else None
            #Find the paragraph where the TOC1 locates
            for m in range(tag.ChildObjects.Count):
                cObj = tag.ChildObjects.get_Item(m)
                if isinstance(cObj, Paragraph):
                    para = cObj if isinstance(cObj, Paragraph) else None
                    if para.StyleName == "TOC2":
                        #Set the tab style of paragraph
                        for n in range(para.Format.Tabs.Count):
                            tab = para.Format.Tabs.get_Item(n)
                            tab.Position = tab.Position + 20
                            tab.TabLeader = TabLeader.NoLeader

#Save the Word file
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()
