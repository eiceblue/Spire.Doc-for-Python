from spire.doc import *
from spire.doc.common import *

outputFile = "RemoveFootnote.docx"
inputFile = "./Data/Footnote.docx"

document = Document()
document.LoadFromFile(inputFile)
section = document.Sections[0]
#traverse paragraphs in the section and find the footnote
for y in range(section.Paragraphs.Count):
    para = section.Paragraphs.get_Item(y)
    index = -1
    i = 0
    cnt = para.ChildObjects.Count
    while i < cnt:
        pBase = para.ChildObjects[i] if isinstance(para.ChildObjects[i], ParagraphBase) else None
        if isinstance(pBase, Footnote):
            index = i
            break
        i += 1
    if index > -1:
        #remove the footnote
        para.ChildObjects.RemoveAt(index)
        
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
