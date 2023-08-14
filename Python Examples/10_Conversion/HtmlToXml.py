from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_HtmlFile.html"
outputFile = "HtmlToXml.xml"
        
#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Save to file.
document.SaveToFile(outputFile, FileFormat.Xml)
document.Close()