from spire.doc import *
from spire.doc.common import *

inputFile_1="./Data/SupportDocumentCompare1.docx"
inputFile_2="./Data/SupportDocumentCompare2.docx"

doc1 = Document()
doc2 = Document()
doc1.LoadFromFile(inputFile_1)
doc2.LoadFromFile(inputFile_2)

doc1.Compare(doc2, "Author")
revisions = DifferRevisions(doc1)
content = ""
m = 0
n = 0
#get insert revision
insertRevisionList = revisions.InsertRevisions

#get deleted revisions
deleteRevisionList = revisions.DeleteRevisions

for i in range(0, insertRevisionList.__len__()):
    # if isinstance(insertRevisionList[i], TextRange):
    if insertRevisionList[i].DocumentObjectType == DocumentObjectType.TextRange:
          m += 1
          textRange = TextRange(insertRevisionList[i])
          content += "insert #" + m.__str__() + ":" + textRange.Text + '\n'   ; content += "=====================" + '\n'
for i in range(0, deleteRevisionList.__len__()):
    # if isinstance(deleteRevisionList[i], TextRange):
    if deleteRevisionList[i].DocumentObjectType == DocumentObjectType.TextRange:
           n += 1
           textRange = TextRange(deleteRevisionList[i])
           content += "delete #" + n.__str__() + ":" + textRange.Text + '\n'      ;  content += "=====================" + '\n'

with open("differRevisions.txt", "w", encoding="utf-8") as file:
    file.write(content)           