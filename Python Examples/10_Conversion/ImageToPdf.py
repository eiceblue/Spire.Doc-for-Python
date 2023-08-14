from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Image.png"
outputFile = "ImageToPdf.pdf"
        
#Create a new document
doc = Document()
#Create a new section
section = doc.AddSection()
#Create a new paragraph
paragraph = section.AddParagraph()
#Add a picture for paragraph
picture = paragraph.AppendPicture(inputFile)
#Set the page size to the same size as picture
#section.PageSetup.PageSize = new SizeF(picture.Width, picture.Height)
#Set A4 page size
section.PageSetup.PageSize = PageSize.A4()
#Set the page margins
section.PageSetup.Margins.Top = 10
section.PageSetup.Margins.Left = 25
doc.SaveToFile(outputFile, FileFormat.PDF)
doc.Close()