from spire.doc import *
from spire.doc.common import *

outputFile = "ChangeTOCStyle.docx"
inputFile =  "./Data/Template_Toc.docx"

#Load document from disk
doc = Document()
doc.LoadFromFile(inputFile)

#Defind a Toc style
tocStyle = Style.CreateBuiltinStyle(BuiltinStyle.Toc1, doc) if isinstance(Style.CreateBuiltinStyle(BuiltinStyle.Toc1, doc), ParagraphStyle) else None
tocStyle.CharacterFormat.FontName = "Aleo"
tocStyle.CharacterFormat.FontSize = 15
tocStyle.CharacterFormat.TextColor = Color.get_CadetBlue()
doc.Styles.Add(tocStyle)

#Loop through sections
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    #Loop through content of section
    for j in range(section.Body.ChildObjects.Count):
        obj = section.Body.ChildObjects.get_Item(j)
        #Find the structure document tag
        if isinstance(obj, StructureDocumentTag):
            tag = obj if isinstance(obj, StructureDocumentTag) else None
            #Find the paragraph where the TOC1 locates
            for k in range(tag.ChildObjects.Count):
                cObj = tag.ChildObjects.get_Item(k)
                if isinstance(cObj, Paragraph):
                    para = cObj if isinstance(cObj, Paragraph) else None
                    if para.StyleName == "TOC1":
                        #Apply the new style for TOC1 paragraph
                        para.ApplyStyle(tocStyle.Name)

#Save the Word file
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()

