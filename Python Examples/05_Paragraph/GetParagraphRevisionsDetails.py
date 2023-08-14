from spire.doc import *
from spire.doc.common import *


def WriteAllText(fname:str,text:List[str]):
        fp = open(fname,"w",encoding='utf-8')
        for s in text:
            fp.write(s)
        fp.close()


inputFile =  "./Data/Revisions.docx"
outputFile = "GetParagraphRevisionsDetails.txt"
       
document = Document()
document.LoadFromFile(inputFile)

builder = ""

#loop paragraph
for i in range(document.Sections.Count):
    section = document.Sections.get_Item(i)
    for j in range(section.Paragraphs.Count):
        paragraph = section.Paragraphs.get_Item(j)
        if paragraph.IsDeleteRevision:
            builder += "The section {} paragraph {} has been changed (deleted).".format(document.GetIndex(section), section.GetIndex(paragraph))
            builder += "\n"
            builder += "Author: "
            builder += paragraph.DeleteRevision.Author
            builder += "\n"
            builder += "DateTime: "
            builder += paragraph.DeleteRevision.DateTime.ToString()
            builder += "\n"
            builder += "Type: " 
            builder += paragraph.DeleteRevision.Type.name
            builder += "\n"
        elif paragraph.IsInsertRevision:
            builder += "The section {} paragraph {} has been changed (inserted).".format(document.GetIndex(section), section.GetIndex(paragraph))
            builder += "Author: " 
            builder += paragraph.InsertRevision.Author
            builder += "\n"
            builder += "DateTime: " 
            builder += paragraph.InsertRevision.DateTime.ToString()
            builder += "\n"
            builder += "Type: " 
            builder += paragraph.InsertRevision.Type.name
            builder += "\n"
        else:
            for k in range(paragraph.ChildObjects.Count):
                obj = paragraph.ChildObjects.get_Item(k)
                if obj.DocumentObjectType is DocumentObjectType.TextRange:
                    textRange = obj if isinstance(obj, TextRange) else None
                    if textRange.IsDeleteRevision:
                        builder += "The section {} paragraph {} textrange {} has been changed (deleted).".format(document.GetIndex(section), section.GetIndex(paragraph), paragraph.GetIndex(textRange))
                        builder += "Author: "
                        builder += textRange.DeleteRevision.Author
                        builder += "\n"
                        builder += "DateTime: " 
                        builder += textRange.DeleteRevision.DateTime.ToString()
                        builder += "\n"
                        builder += "Type: " 
                        builder += textRange.DeleteRevision.Type.name
                        builder += "\n"
                        builder += "Change Text: " 
                        builder += textRange.Text
                        builder += "\n"
                    elif textRange.IsInsertRevision:
                        builder += "The section {} paragraph {} textrange {} has been changed (inserted).".format(document.GetIndex(section), section.GetIndex(paragraph), paragraph.GetIndex(textRange))
                        builder += "Author: " 
                        builder += textRange.InsertRevision.Author
                        builder += "\n"
                        builder += "DateTime: " 
                        builder += textRange.InsertRevision.DateTime.ToString()
                        builder += "\n"
                        builder += "Type: " 
                        builder += textRange.InsertRevision.Type.name
                        builder += "\n"
                        builder += "Change Text: " 
                        builder += textRange.Text
                        builder += "\n"

#Write the contents in a TXT file
WriteAllText(outputFile, builder)
