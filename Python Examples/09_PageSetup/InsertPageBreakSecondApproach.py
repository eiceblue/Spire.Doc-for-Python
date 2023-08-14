from spire.doc import *
from spire.doc.common import *

inputFile =  "./Data/Template_Docx_1.docx"
outputFile = "InsertPageBreakSecondApproach.docx"
#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Insert page break.
document.Sections[0].Paragraphs[3].AppendBreak(BreakType.PageBreak)
#Save the file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()

