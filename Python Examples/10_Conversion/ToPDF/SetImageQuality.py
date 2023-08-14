
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Doc_1.doc"
outputFile =  "SetImageQuality.pdf"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile, FileFormat.Doc)
#Set the output image quality to be 40% of the original image. The default set of the output image quality is 80% of the original.
document.JPEGQuality = 40
#Save to file.
document.SaveToFile(outputFile, FileFormat.PDF)
document.Close()
