from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Sample.docx"
outputFile = "ReplaceWithText.docx"

#Create word document
document = Document()

#Load the document from disk.
document.LoadFromFile(inputFile)

#Replace text
document.Replace("word", "ReplacedText", False, True)

#Save doc file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()