from spire.doc import *
from spire.doc.common import *


outputFile = "ASCIICharactersBulletStyle.docx"

        
#Create a new document
document = Document()
section = document.AddSection()

#Create four list styles based on different ASCII characters
listStyle1 = ListStyle(document, ListType.Bulleted)
listStyle1.Name = "liststyle"
listStyle1.Levels[0].BulletCharacter = "\u006e"
listStyle1.Levels[0].CharacterFormat.FontName = "Wingdings"
document.ListStyles.Add(listStyle1)
listStyle2 = ListStyle(document, ListType.Bulleted)
listStyle2.Name = "liststyle2"
listStyle2.Levels[0].BulletCharacter = "\u0075"
listStyle2.Levels[0].CharacterFormat.FontName = "Wingdings"
document.ListStyles.Add(listStyle2)
listStyle3 = ListStyle(document, ListType.Bulleted)
listStyle3.Name = "liststyle3"
listStyle3.Levels[0].BulletCharacter = "\u00b2"
listStyle3.Levels[0].CharacterFormat.FontName = "Wingdings"
document.ListStyles.Add(listStyle3)
listStyle4 = ListStyle(document, ListType.Bulleted)
listStyle4.Name = "liststyle4"
listStyle4.Levels[0].BulletCharacter = "\u00d8"
listStyle4.Levels[0].CharacterFormat.FontName = "Wingdings"
document.ListStyles.Add(listStyle4)

#Add four paragraphs and apply list style separately
p1 = section.Body.AddParagraph()
p1.AppendText("Spire.Doc for .NET")
p1.ListFormat.ApplyStyle(listStyle1.Name)
p2 = section.Body.AddParagraph()
p2.AppendText("Spire.Doc for Java")
p2.ListFormat.ApplyStyle(listStyle2.Name)
p3 = section.Body.AddParagraph()
p3.AppendText("Spire.Doc for C++")
p3.ListFormat.ApplyStyle(listStyle3.Name)
p4 = section.Body.AddParagraph()
p4.AppendText("Spire.Doc for Python")
p4.ListFormat.ApplyStyle(listStyle4.Name)

#Save to docx file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()