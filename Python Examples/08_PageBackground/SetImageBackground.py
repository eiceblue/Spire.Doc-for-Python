from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_2.docx"
inputFile_Img = "./Data/Background.png"
outputFile = "SetImageBackground.docx"

#load a word document
document = Document()
document.LoadFromFile(inputFile)
#set the background type as picture.
document.Background.Type = BackgroundType.Picture
#set the background picture
document.Background.SetPicture(inputFile_Img)
#save the file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()