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
listSty2 = ListStyle(document, ListType.Numbered)
for i in range(listSty2.Levels.Count):
    listLev = listSty2.Levels.get_Item(i)
    listLev.UsePrevLevelPattern = True
    listLev.NumberPrefix = "1."
listSty2.Name = "MyStyle2"
document.ListStyles.Add(listSty2)
paragraph.ListFormat.ApplyStyle(listSty2.Name)

#Add list style 3.
listSty3 = ListStyle(document, ListType.Numbered)
for i in range(listSty3.Levels.Count):
    listLev = listSty3.Levels.get_Item(i)
    listLev.UsePrevLevelPattern = True
    listLev.NumberPrefix = "1.1."
listSty3.Name = "MyStyle3"
document.ListStyles.Add(listSty3)

#Add Heading 3.
for i in range(0, 4):
    paragraph = section.AddParagraph()

    #Append text
    paragraph.AppendText(BuiltinStyle.Heading3.name)

    #Apply list style 3 for Heading 3
    paragraph.ApplyStyle(BuiltinStyle.Heading3)
    paragraph.ListFormat.ApplyStyle(listSty3.Name)
#Save the file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
