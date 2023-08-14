from spire.doc import *
from spire.doc.common import *

inputFile =  "./Data/Template_RtfFile.rtf"
outputFile = "RtfToHtml.html"
               
#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Save to file.
document.SaveToFile(outputFile, FileFormat.Html)
document.Close()
