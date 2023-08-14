from spire.doc import *
from spire.doc.common import *

outputFile = "AddContentControls.docx"

#Creat a new word document.
document = Document()
section = document.AddSection()
paragraph = section.AddParagraph()
txtRange = paragraph.AppendText("The following example shows how to add content controls in a Word document.")
paragraph = section.AddParagraph()

#Add Combo Box Content Control.
paragraph = section.AddParagraph()
txtRange = paragraph.AppendText("Add Combo Box Content Control:  ")
txtRange.CharacterFormat.Italic = True
sd = StructureDocumentTagInline(document)
paragraph.ChildObjects.Add(sd)
sd.SDTProperties.SDTType = SdtType.ComboBox
cb = SdtComboBox()
cb.ListItems.Add(SdtListItem("Spire.Doc"))
cb.ListItems.Add(SdtListItem("Spire.XLS"))
cb.ListItems.Add(SdtListItem("Spire.PDF"))
sd.SDTProperties.ControlProperties = cb
rt = TextRange(document)
rt.Text = cb.ListItems[0].DisplayText
sd.SDTContent.ChildObjects.Add(rt)
section.AddParagraph()

#Add Text Content Control.
paragraph = section.AddParagraph()
txtRange = paragraph.AppendText("Add Text Content Control:  ")
txtRange.CharacterFormat.Italic = True
sd = StructureDocumentTagInline(document)
paragraph.ChildObjects.Add(sd)
sd.SDTProperties.SDTType = SdtType.Text
text = SdtText(True)
text.IsMultiline = True
sd.SDTProperties.ControlProperties = text
rt = TextRange(document)
rt.Text = "Text"
sd.SDTContent.ChildObjects.Add(rt)
section.AddParagraph()

#Add Picture Content Control.
paragraph = section.AddParagraph()
txtRange = paragraph.AppendText("Add Picture Content Control:  ")
txtRange.CharacterFormat.Italic = True
sd = StructureDocumentTagInline(document)
paragraph.ChildObjects.Add(sd)
sd.SDTProperties.SDTType = SdtType.Picture
pic = DocPicture(document)
pic.Width = 10
pic.Height = 10
pic.LoadImage("./Data/logo.png")
sd.SDTContent.ChildObjects.Add(pic)
section.AddParagraph()

#Add Date Picker Content Control.
paragraph = section.AddParagraph()
txtRange = paragraph.AppendText("Add Date Picker Content Control:  ")
txtRange.CharacterFormat.Italic = True
sd = StructureDocumentTagInline(document)
paragraph.ChildObjects.Add(sd)
sd.SDTProperties.SDTType = SdtType.DatePicker
date = SdtDate()
date.CalendarType = CalendarType.Default
date.DateFormat = "yyyy.MM.dd"
date.FullDate = DateTime.get_Now()
sd.SDTProperties.ControlProperties = date
rt = TextRange(document)
rt.Text = "1990.02.08"
sd.SDTContent.ChildObjects.Add(rt)
section.AddParagraph()

#Add Drop-Down List Content Control.
paragraph = section.AddParagraph()
txtRange = paragraph.AppendText("Add Drop-Down List Content Control:  ")
txtRange.CharacterFormat.Italic = True
sd = StructureDocumentTagInline(document)
paragraph.ChildObjects.Add(sd)
sd.SDTProperties.SDTType = SdtType.DropDownList
sddl = SdtDropDownList()
sddl.ListItems.Add(SdtListItem("Harry"))
sddl.ListItems.Add(SdtListItem("Jerry"))
sd.SDTProperties.ControlProperties = sddl
rt = TextRange(document)
rt.Text = sddl.ListItems[0].DisplayText
sd.SDTContent.ChildObjects.Add(rt)

#Save the document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()