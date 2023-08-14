from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_1.docx"
outputFile =  "InsertSectionBreak.docx"
#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Insert section break. There are five section break options including EvenPage, NewColumn, NewPage, NoBreak, OddPage.
document.Sections[0].Paragraphs[1].InsertSectionBreak(SectionBreakType.NoBreak)
#Save the file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()

