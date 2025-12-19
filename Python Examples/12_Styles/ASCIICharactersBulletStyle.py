from spire.doc import *
from spire.doc.common import *


outputFile = "ASCIICharactersBulletStyle.docx"

        
#Create a new document
document = Document()
section = document.AddSection()

#Create four list styles based on different ASCII characters
listStyle1 =  document.Styles.Add(ListType.Bulleted, "liststyle")
Levels = listStyle1.ListRef.Levels
Levels[0].BulletCharacter = "\u006e"
Levels[0].CharacterFormat.FontName = "Wingdings"

listStyle2 = document.Styles.Add(ListType.Bulleted, "liststyle2")
Levels2 = listStyle2.ListRef.Levels

Levels2[0].BulletCharacter = "\u0075"
Levels2[0].CharacterFormat.FontName = "Wingdings"

listStyle3 = document.Styles.Add(ListType.Bulleted, "liststyle3")
Levels3 = listStyle3.ListRef.Levels
Levels3[0].BulletCharacter = "\u00b2"
Levels3[0].CharacterFormat.FontName = "Wingdings"

listStyle4 = document.Styles.Add(ListType.Bulleted, "liststyle4")
Levels4 = listStyle4.ListRef.Levels
Levels4[0].BulletCharacter = "\u00d8"
Levels4[0].CharacterFormat.FontName = "Wingdings"

#Add four paragraphs and apply list style separately
p1 = section.Body.AddParagraph()
p1.AppendText("Spire.Doc for .NET")
p1.ListFormat.ApplyStyle(listStyle1.Name)
p2 = section.Body.AddParagraph()
p2.AppendText("Spire.Doc for .NET")
p2.ListFormat.ApplyStyle(listStyle2.Name)
p3 = section.Body.AddParagraph()
p3.AppendText("Spire.Doc for .NET")
p3.ListFormat.ApplyStyle(listStyle3.Name)
p4 = section.Body.AddParagraph()
p4.AppendText("Spire.Doc for .NET")
p4.ListFormat.ApplyStyle(listStyle4.Name)

#Save to docx file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()