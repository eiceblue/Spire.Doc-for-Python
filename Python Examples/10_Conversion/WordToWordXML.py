from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_1.docx"
outputFile_2003 = "WordToWordML.xml"
outputFile_2007 = "WordToWordXML.xml"
      

#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)

#For word 2003:
document.SaveToFile(outputFile_2003, FileFormat.WordML)

#For word 2007:
document.SaveToFile(outputFile_2007, FileFormat.WordXml)
document.Close()

