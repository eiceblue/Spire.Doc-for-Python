from spire.doc import *
from spire.doc.common import *

outputFile = "AddHorizontalLine.docx"
#Create Word document.
doc = Document()
sec = doc.AddSection()
para = sec.AddParagraph()
para.AppendHorizonalLine()
#Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()

