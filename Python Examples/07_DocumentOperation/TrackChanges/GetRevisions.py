
from spire.doc import *
from spire.doc.common import *

def WriteAllText(fname:str,text:List[str]):
        fp = open(fname,"w")
        for s in text:
            fp.write(s)
        fp.close()

inputFile = "./Data/GetRevisions.docx"
outputFile1 = "insertRevisions.txt"
outputFile2 = "deleteRevisions.txt"

#Create Word document.
document = Document()
document.LoadFromFile(inputFile)
insertRevision = ""
insertRevision += "Insert revisions:"
insertRevision += "\n"
index_insertRevision = 0
deleteRevision = ""
deleteRevision += "Delete revisions:"
deleteRevision += "\n"
index_deleteRevision = 0
#Traverse sections
for k in range(document.Sections.Count):
    sec = document.Sections.get_Item(k)
    #Iterate through the element under body in the section
    for m in range(sec.Body.ChildObjects.Count):
        docItem = sec.Body.ChildObjects.get_Item(m)
        if isinstance(docItem, Paragraph):
            para = docItem
            #Determine if the paragraph is an insertion revision
            if para.IsInsertRevision:
                index_insertRevision += 1
                insertRevision += "Index: " 
                insertRevision += str(index_insertRevision)
                insertRevision += "\n"
                #Get insertion revision
                insRevison = para.InsertRevision
                #Get insertion revision type
                insType = insRevison.Type
                insertRevision += "Type: "
                insertRevision += insType.name
                insertRevision += "\n"
                #Get insertion revision author
                insAuthor = insRevison.Author
                insertRevision += "Author: "
                insertRevision += insAuthor
                insertRevision += "\n"
            #Determine if the paragraph is a delete revision
            elif para.IsDeleteRevision:
                index_deleteRevision += 1
                deleteRevision += "Index: "
                deleteRevision += str(index_deleteRevision)
                deleteRevision += "\n"
                delRevison = para.DeleteRevision
                delType = delRevison.Type
                deleteRevision += "Type: "
                deleteRevision += delType.name
                deleteRevision += "\n"
                delAuthor = delRevison.Author
                deleteRevision += "Author: "
                deleteRevision += delAuthor
                deleteRevision += "\n"
            #Iterate through the element in the paragraph
            for j in range(para.ChildObjects.Count):
                obj = para.ChildObjects.get_Item(j)
                if isinstance(obj, TextRange):
                    textRange = obj
                    #Determine if the textrange is an insertion revision
                    if textRange.IsInsertRevision:
                        index_insertRevision += 1
                        insertRevision += "Index: "
                        insertRevision += str(index_insertRevision)
                        insertRevision += "\n"
                        insRevison = textRange.InsertRevision
                        insType = insRevison.Type
                        insertRevision += "Type: "
                        insertRevision += insType.name
                        insertRevision += "\n"
                        insAuthor = insRevison.Author
                        insertRevision += "Author: " 
                        insertRevision += insAuthor
                        insertRevision += "\n"
                    elif textRange.IsDeleteRevision:
                        index_deleteRevision += 1
                        deleteRevision += "Index: "
                        deleteRevision += str(index_deleteRevision)
                        deleteRevision += "\n"
                        #Determine if the textrange is a delete revision
                        delRevison = textRange.DeleteRevision
                        delType = delRevison.Type
                        deleteRevision += "Type: "
                        deleteRevision += delType.name
                        deleteRevision += "\n"
                        delAuthor = delRevison.Author
                        deleteRevision += "Author: "
                        deleteRevision += delAuthor
                        deleteRevision += "\n"
WriteAllText(outputFile1, insertRevision)
WriteAllText(outputFile2, deleteRevision)