
from spire.doc import *
from spire.doc.common import *



inputFile = "./Data/Summary_of_Science.doc"
outputFile = "ToXML.xml"
#Create word document.
document = Document()
document.LoadFromFile(inputFile)
#Save the document to a xml file.
document.SaveToFile(outputFile, FileFormat.Xml)
document.Close()