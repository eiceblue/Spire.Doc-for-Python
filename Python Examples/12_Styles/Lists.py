from spire.doc import *
from spire.doc.common import *


outputFile = "Lists.docx"


#Initialize a document
document = Document()
#Add a section
sec = document.AddSection()
#Add paragraph and set list style
paragraph = sec.AddParagraph()
paragraph.AppendText("Lists")
paragraph.ApplyStyle(BuiltinStyle.Title)

paragraph = sec.AddParagraph()
paragraph.AppendText("Numbered List:").CharacterFormat.Bold = True

#Create list style
numberList = ListStyle(document, ListType.Numbered)
numberList.Name = "numberList"
#%1-%9
numberList.Levels[1].NumberPrefix = "%1."
numberList.Levels[1].PatternType = ListPatternType.Arabic
numberList.Levels[2].NumberPrefix = "%1.%2."
numberList.Levels[2].PatternType = ListPatternType.Arabic

bulletList = ListStyle(document, ListType.Bulleted)
bulletList.Name = "bulletList"

#add the list style into document
document.ListStyles.Add(numberList)
document.ListStyles.Add(bulletList)

#Add paragraph and apply the list style
paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 1")
paragraph.ListFormat.ApplyStyle(numberList.Name)

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2")
paragraph.ListFormat.ApplyStyle(numberList.Name)

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.1")
paragraph.ListFormat.ApplyStyle(numberList.Name)
paragraph.ListFormat.ListLevelNumber = 1

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.2")
paragraph.ListFormat.ApplyStyle(numberList.Name)
paragraph.ListFormat.ListLevelNumber = 1

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.2.1")
paragraph.ListFormat.ApplyStyle(numberList.Name)
paragraph.ListFormat.ListLevelNumber = 2
paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.2.2")
paragraph.ListFormat.ApplyStyle(numberList.Name)
paragraph.ListFormat.ListLevelNumber = 2
paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.2.3")
paragraph.ListFormat.ApplyStyle(numberList.Name)
paragraph.ListFormat.ListLevelNumber = 2

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.3")
paragraph.ListFormat.ApplyStyle(numberList.Name)
paragraph.ListFormat.ListLevelNumber = 1

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 3")
paragraph.ListFormat.ApplyStyle(numberList.Name)

paragraph = sec.AddParagraph()
paragraph.AppendText("Bulleted List:").CharacterFormat.Bold = True

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 1")
paragraph.ListFormat.ApplyStyle(bulletList.Name)
paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2")
paragraph.ListFormat.ApplyStyle(bulletList.Name)

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.1")
paragraph.ListFormat.ApplyStyle(bulletList.Name)
paragraph.ListFormat.ListLevelNumber = 1
paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.2")
paragraph.ListFormat.ApplyStyle(bulletList.Name)
paragraph.ListFormat.ListLevelNumber = 1
paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 3")
paragraph.ListFormat.ApplyStyle(bulletList.Name)

#Save doc file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()