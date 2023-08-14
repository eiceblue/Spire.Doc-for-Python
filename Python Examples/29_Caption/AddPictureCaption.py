from spire.doc import *
from spire.doc.common import *


inputFile1 = "./Data/Spire.Doc.png"
inputFile2 = "./Data/Word.png"
outputFile = "AddPictureCaption.docx"

#Create word document
document = Document()

#Create a new section
section = document.AddSection()

#Add the first picture
par1 = section.AddParagraph()
par1.Format.AfterSpacing = 10.0
pic1 = par1.AppendPicture(inputFile1)

pic1.Height = 100.0
pic1.Width = 120.0
#Add caption to the picture
tempFormat = CaptionNumberingFormat.Number
pic1.AddCaption("Figure", tempFormat, CaptionPosition.BelowItem)

#Add the second picture
par2 = section.AddParagraph()
pic2 = par2.AppendPicture(inputFile2)

pic2.Height = 100.0
pic2.Width = 120.0
#Add caption to the picture
pic2.AddCaption("Figure", tempFormat, CaptionPosition.BelowItem)

#Update fields
document.IsUpdateFields = True

#Save the Word document
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()