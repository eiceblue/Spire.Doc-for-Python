from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/AcceptOrRejectTrackedChanges.docx"
outputFile = "AcceptOrRejectTrackedChanges_out.docx"


#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Get the first section and the paragraph we want to accept/reject the changes.
sec = document.Sections[0]
para = sec.Paragraphs[0]
#Accept the changes or reject the changes.
para.Document.AcceptChanges()
#para.Document.RejectChanges()
#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()