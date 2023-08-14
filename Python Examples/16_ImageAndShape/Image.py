from spire.doc import *
from spire.doc.common import *


outputFile = "Image.docx"
def _InsertImage(section):
    #Add paragraph
    paragraph = section.AddParagraph()
    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Left
    picture = paragraph.AppendPicture("./Data/Spire.Doc.png")

    picture.Width = 100
    picture.Height = 100

    paragraph = section.AddParagraph()
    paragraph.Format.LineSpacing = 20
    tr = paragraph.AppendText("Spire.Doc for .NET is a professional Word .NET library specially designed for developers to create, read, write, convert and print Word document files from any .NET( C#, VB.NET, ASP.NET) platform with fast and high quality performance. ")
    tr.CharacterFormat.FontName = "Arial"
    tr.CharacterFormat.FontSize = 14

    section.AddParagraph()
    paragraph = section.AddParagraph()
    paragraph.Format.LineSpacing = 20
    tr = paragraph.AppendText("As an independent Word .NET component, Spire.Doc for .NET doesn't need Microsoft Word to be installed on the machine. However, it can incorporate Microsoft Word document creation capabilities into any developers' .NET applications.")
    tr.CharacterFormat.FontName = "Arial"
    tr.CharacterFormat.FontSize = 14



#Create a document
document = Document()

#Add a seciton
section = document.AddSection()

#insert image
_InsertImage(section)

#Save doc file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()

