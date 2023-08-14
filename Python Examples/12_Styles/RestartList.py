from spire.doc import *
from spire.doc.common import *

outputFile = "RestartList.docx"
#Create word document
document = Document()

#Create a new section
section = document.AddSection()

#Create a new paragraph
paragraph = section.AddParagraph()

#Append Text
paragraph.AppendText("List 1")

numberList = ListStyle(document, ListType.Numbered)
numberList.Name = "Numbered1"
document.ListStyles.Add(numberList)

#Add paragraph and apply the list style
paragraph = section.AddParagraph()
paragraph.AppendText("List Item 1")
paragraph.ListFormat.ApplyStyle(numberList.Name)

paragraph = section.AddParagraph()
paragraph.AppendText("List Item 2")
paragraph.ListFormat.ApplyStyle(numberList.Name)

paragraph = section.AddParagraph()
paragraph.AppendText("List Item 3")
paragraph.ListFormat.ApplyStyle(numberList.Name)

paragraph = section.AddParagraph()
paragraph.AppendText("List Item 4")
paragraph.ListFormat.ApplyStyle(numberList.Name)

#Append Text
paragraph = section.AddParagraph()
paragraph.AppendText("List 2")

numberList2 = ListStyle(document, ListType.Numbered)
numberList2.Name = "Numbered2"
#set start number of second list
numberList2.Levels[0].StartAt = 10
document.ListStyles.Add(numberList2)

#Add paragraph and apply the list style
paragraph = section.AddParagraph()
paragraph.AppendText("List Item 5")
paragraph.ListFormat.ApplyStyle(numberList2.Name)

paragraph = section.AddParagraph()
paragraph.AppendText("List Item 6")
paragraph.ListFormat.ApplyStyle(numberList2.Name)

paragraph = section.AddParagraph()
paragraph.AppendText("List Item 7")
paragraph.ListFormat.ApplyStyle(numberList2.Name)

paragraph = section.AddParagraph()
paragraph.AppendText("List Item 8")
paragraph.ListFormat.ApplyStyle(numberList2.Name)

#Save to docx file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
