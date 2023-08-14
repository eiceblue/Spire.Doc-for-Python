from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Template_Docx_5.docx"
outputFile = "CopyParagraph.docx"

#Create Word document1.
document1 = Document()

#Load the file from disk.
document1.LoadFromFile(inputFile)

#Create a new document.
document2 = Document()

#Get paragraph 1 and paragraph 2 in document1.
s = document1.Sections[0]
p1 = s.Paragraphs[0]
p2 = s.Paragraphs[1]

#Copy p1 and p2 to document2.
s2 = document2.AddSection()
NewPara1 = p1.Clone()
s2.Paragraphs.Add(NewPara1)
NewPara2 = p2.Clone()
s2.Paragraphs.Add(NewPara2)

#Add watermark.
WM = PictureWatermark()
WM.SetPicture("./Data/Logo.jpg")
document2.Watermark = WM

#Save the file.
document2.SaveToFile(outputFile, FileFormat.Docx2013)
document2.Close()

