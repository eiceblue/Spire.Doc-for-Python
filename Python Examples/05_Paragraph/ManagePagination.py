from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Template_Docx_1.docx"
outputFile = "ManagePagination.docx"

#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)

#Get the first section and the paragraph we want to manage the pagination.
sec = document.Sections[0]
para = sec.Paragraphs[4]

#Set the pagination format as Format.PageBreakBefore for the checked paragraph.
para.Format.PageBreakBefore = True

#Save the file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
