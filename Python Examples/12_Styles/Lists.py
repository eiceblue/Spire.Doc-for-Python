from spire.doc import *
from spire.doc.common import *


outputFile = "Lists.docx"


#Initialize a document
document =  Document()
#Add a section
sec = document.AddSection()
#Add paragraph and set list style
paragraph = sec.AddParagraph()
paragraph.AppendText("Lists")
paragraph.ApplyStyle(BuiltinStyle.Title)

paragraph = sec.AddParagraph()
paragraph.AppendText("Numbered List:").CharacterFormat.Bold = True

#Create list style
numberList = document.Styles.Add(ListType.Numbered, "numberList")
Levels = numberList.ListRef.Levels
#%1-%9
Levels[1].NumberPrefix = "%1."
Levels[1].PatternType = ListPatternType.Arabic
Levels[2].NumberPrefix = "%1.%2."
Levels[2].PatternType = ListPatternType.Arabic

bulletList = document.Styles.Add(ListType.Bulleted, "bulletList")

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
paragraph.ListFormat.ApplyStyle(bulletList)
paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2")
paragraph.ListFormat.ApplyStyle(bulletList)

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.1")
paragraph.ListFormat.ApplyStyle(bulletList)
paragraph.ListFormat.ListLevelNumber = 1
paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.2")
paragraph.ListFormat.ApplyStyle(bulletList)
paragraph.ListFormat.ListLevelNumber = 1
paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 3")
paragraph.ListFormat.ApplyStyle(bulletList)

#Save doc file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()