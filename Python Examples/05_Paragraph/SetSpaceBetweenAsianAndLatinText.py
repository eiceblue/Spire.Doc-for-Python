from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/SetSpaceBetweenAsianAndLatinText.docx"
outputFile = "SetSpaceBetweenAsianAndLatinText.docx"

document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)

para = document.Sections[0].Paragraphs[0]

#Set whether to automatically adjust space between Asian text and Latin text
para.Format.AutoSpaceDE = False
#Set whether to automatically adjust space between Asian text and numbers
para.Format.AutoSpaceDN = True

#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()


