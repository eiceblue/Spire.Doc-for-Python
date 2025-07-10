from spire.doc import *
from spire.doc.common import *

outputFile = "XMLToWord.docx"
inputFile = "Data/Template_XmlFile.xml"

#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile, FileFormat.Xml)

#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
