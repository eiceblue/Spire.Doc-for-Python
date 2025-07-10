from spire.doc import *
from spire.doc.common import *

outputFile = "XMLToPDF.pdf"
inputFile = "Data/Template_XmlFile.xml"

#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile, FileFormat.Xml)

#Save to file.
document.SaveToFile(outputFile, FileFormat.PDF)
document.Close()
