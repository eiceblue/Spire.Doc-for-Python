from spire.doc import *
from spire.doc.common import *


outputFile = "FormACatalogue.docx"
 
#Create Word document.
document = Document()

#Add a new section. 
section = document.AddSection()
paragraph = section.Paragraphs[0] if section.Paragraphs.Count > 0 else section.AddParagraph()

#Add Heading 1.
paragraph = section.AddParagraph()
paragraph.AppendText(BuiltinStyle.Heading1.name)
paragraph.ApplyStyle(BuiltinStyle.Heading1)
paragraph.ListFormat.ApplyNumberedStyle()

#Add Heading 2.
paragraph = section.AddParagraph()
paragraph.AppendText(BuiltinStyle.Heading2.name)
paragraph.ApplyStyle(BuiltinStyle.Heading2)

#List style for Headings 2.
listStyle2 =  document.Styles.Add(ListType.Numbered, "MyStyle2")
Levels = listStyle2.ListRef.Levels
for i in range(Levels.Count):
    listLev = Levels.get_Item(i)
    listLev.UsePrevLevelPattern = True
    listLev.NumberPrefix = "1."
paragraph.ListFormat.ApplyStyle(listStyle2.Name)

#Add list style 3.
listStyle3 = document.Styles.Add(ListType.Numbered, "MyStyle3")
Levels1 = listStyle3.ListRef.Levels
for i in range(Levels1.Count):
    listLev = Levels1.get_Item(i)
    listLev.UsePrevLevelPattern = True
    listLev.NumberPrefix = "1.1."

#Add Heading 3.
for i in range(0, 4):
    paragraph = section.AddParagraph()

    #Append text
    paragraph.AppendText(BuiltinStyle.Heading3.name)

    #Apply list style 3 for Heading 3
    paragraph.ApplyStyle(BuiltinStyle.Heading3)
    paragraph.ListFormat.ApplyStyle(listStyle3.Name)
#Save the file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()

