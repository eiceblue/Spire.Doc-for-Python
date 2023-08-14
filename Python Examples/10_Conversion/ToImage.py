from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ConvertedTemplate.docx"
outputFile =  "ToImage.png"
#Create word document
document = Document()
document.LoadFromFile(inputFile)
#Obtain image data in the default format of png,you can use it to convert other image format.
imageStream = document.SaveImageToStreams(0, ImageType.Bitmap)
with open(outputFile,'wb') as imageFile:
    imageFile.write(imageStream.ToArray())
document.Close()
