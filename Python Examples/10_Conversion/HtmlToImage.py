from spire.doc import *
from spire.doc.common import *
import io

inputFile =  "./Data/Template_HtmlFile1.html"
outputFile = "HtmlToImage.png"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile, FileFormat.Html, XHTMLValidationType.none)
#Obtain image data in the default format of png,you can use it to convert other image format.
imageStream = document.SaveImageToStreams(0, ImageType.Bitmap)
with open(outputFile,'wb') as imageFile:
    imageFile.write(imageStream.ToArray())
document.Close()
