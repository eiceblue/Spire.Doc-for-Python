from spire.doc import *
from spire.doc.common import *


outputFile = "Styles.docx"     
#Initialize a document
document =Document()
sec = document.AddSection()
#Add default title style to document and modify
titleStyle = document.AddStyle(BuiltinStyle.Title)
titleStyle.CharacterFormat.FontName= "cambria"
titleStyle.CharacterFormat.FontSize = 28
titleStyle.CharacterFormat.TextColor = Color.FromArgb(255,42, 123, 136)
#judge if it is Paragraph Style and then set paragraph format
if isinstance(titleStyle, ParagraphStyle):
    ps = titleStyle if isinstance(titleStyle, ParagraphStyle) else None
    ps.ParagraphFormat.Borders.Bottom.BorderType = BorderStyle.Single
    ps.ParagraphFormat.Borders.Bottom.Color = Color.FromArgb(255,42, 123, 136)
    ps.ParagraphFormat.Borders.Bottom.LineWidth = 1.5
    ps.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left
#Add default normal style and modify
normalStyle = document.AddStyle(BuiltinStyle.Normal)
normalStyle.CharacterFormat.FontName = "cambria"
normalStyle.CharacterFormat.FontSize = 11

#Add default heading1 style
heading1Style = document.AddStyle(BuiltinStyle.Heading1)

heading1Style.CharacterFormat.FontName = "cambria"
heading1Style.CharacterFormat.FontSize = 14

heading1Style.CharacterFormat.Bold = True
heading1Style.CharacterFormat.TextColor = Color.FromArgb(255,42, 123, 136)
#Add default heading2 style
heading2Style = document.AddStyle(BuiltinStyle.Heading2)
heading2Style.CharacterFormat.FontName = "cambria"
heading2Style.CharacterFormat.FontSize = 12
heading2Style.CharacterFormat.Bold = True

#List style
bulletList = ListStyle(document, ListType.Bulleted)
bulletList.CharacterFormat.FontName = "cambria"
bulletList.CharacterFormat.FontSize = 12
bulletList.Name = "bulletList"
document.ListStyles.Add(bulletList)

#Apply the style
paragraph = sec.AddParagraph()
paragraph.AppendText("Your Name")
paragraph.ApplyStyle(BuiltinStyle.Title)

paragraph = sec.AddParagraph()
paragraph.AppendText("Address, City, ST ZIP Code | Telephone | Email")
paragraph.ApplyStyle(BuiltinStyle.Normal)

paragraph = sec.AddParagraph()
paragraph.AppendText("Objective")
paragraph.ApplyStyle(BuiltinStyle.Heading1)

paragraph = sec.AddParagraph()
paragraph.AppendText("To get started right away, just click any placeholder text (such as this) and start typing to replace it with your own.")
paragraph.ApplyStyle(BuiltinStyle.Normal)

paragraph = sec.AddParagraph()
paragraph.AppendText("Education")
paragraph.ApplyStyle(BuiltinStyle.Heading1)

paragraph = sec.AddParagraph()
paragraph.AppendText("DEGREE | DATE EARNED | SCHOOL")
paragraph.ApplyStyle(BuiltinStyle.Heading2)

paragraph = sec.AddParagraph()
paragraph.AppendText("Major:Text")
paragraph.ListFormat.ApplyStyle("bulletList")
paragraph = sec.AddParagraph()
paragraph.AppendText("Minor:Text")
paragraph.ListFormat.ApplyStyle("bulletList")
paragraph = sec.AddParagraph()
paragraph.AppendText("Related coursework:Text")
paragraph.ListFormat.ApplyStyle("bulletList")

paragraph = sec.AddParagraph()
paragraph.AppendText("Skills & Abilities")
paragraph.ApplyStyle(BuiltinStyle.Heading1)

paragraph = sec.AddParagraph()
paragraph.AppendText("MANAGEMENT")
paragraph.ApplyStyle(BuiltinStyle.Heading2)

paragraph = sec.AddParagraph()
paragraph.AppendText("Think a document that looks this good has to be difficult to format? Think again! To easily apply any text formatting you see in this document with just a click, on the Home tab of the ribbon, check out Styles.")
paragraph.ListFormat.ApplyStyle("bulletList")

paragraph = sec.AddParagraph()
paragraph.AppendText("COMMUNICATION")
paragraph.ApplyStyle(BuiltinStyle.Heading2)

paragraph = sec.AddParagraph()
paragraph.AppendText("You delivered that big presentation to rave reviews. Don’t be shy about it now! This is the place to show how well you work and play with others.")
paragraph.ListFormat.ApplyStyle("bulletList")

paragraph = sec.AddParagraph()
paragraph.AppendText("LEADERSHIP")
paragraph.ApplyStyle(BuiltinStyle.Heading2)

paragraph = sec.AddParagraph()
paragraph.AppendText("Are you president of your fraternity, head of the condo board, or a team lead for your favorite charity? You’re a natural leader—tell it like it is!")
paragraph.ListFormat.ApplyStyle("bulletList")

paragraph = sec.AddParagraph()
paragraph.AppendText("Experience")
paragraph.ApplyStyle(BuiltinStyle.Heading1)

paragraph = sec.AddParagraph()
paragraph.AppendText("JOB TITLE | COMPANY | DATES FROM - TO")
paragraph.ApplyStyle(BuiltinStyle.Heading2)

paragraph = sec.AddParagraph()
paragraph.AppendText("This is the place for a brief summary of your key responsibilities and most stellar accomplishments.")
paragraph.ListFormat.ApplyStyle("bulletList")

#Save to docx file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()

