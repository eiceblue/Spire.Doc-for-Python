
from spire.doc import *
from spire.doc.common import *


outputFile = "InsertBreak.docx"
#Create word document
document = Document()
section = document.AddSection()

#page setup
section.PageSetup.PageSize = PageSize.A4()
section.PageSetup.Margins.Top = 72
section.PageSetup.Margins.Bottom = 72
section.PageSetup.Margins.Left = 89.85
section.PageSetup.Margins.Right = 89.85

#Add cover.
small = ParagraphStyle(section.Document)
small.Name = "small"
small.CharacterFormat.FontName = "Arial"
small.CharacterFormat.FontSize = 9
small.CharacterFormat.TextColor = Color.get_Gray()
section.Document.Styles.Add(small)
paragraph = section.AddParagraph()
paragraph.AppendText("The sample demonstrates how to insert section break.")
paragraph.ApplyStyle(small.Name)
title = section.AddParagraph()
text = title.AppendText("Field Types Supported by Spire.Doc")
text.CharacterFormat.FontName = "Arial"
text.CharacterFormat.FontSize = 20
text.CharacterFormat.Bold = True
title.Format.BeforeSpacing = section.PageSetup.PageSize.Height / 2 - 3 * section.PageSetup.Margins.Top
title.Format.AfterSpacing = 8
title.Format.HorizontalAlignment = HorizontalAlignment.Right
paragraph = section.AddParagraph()
paragraph.AppendText("e-iceblue Spire.Doc team.")
paragraph.ApplyStyle(small.Name)
paragraph.Format.HorizontalAlignment = HorizontalAlignment.Right

#insert a break code
section = document.AddSection()
section.AddParagraph().InsertSectionBreak(SectionBreakType.NewPage)

#add content
list = ParagraphStyle(section.Document)
list.Name = "list"
list.CharacterFormat.FontName = "Arial"
list.CharacterFormat.FontSize = 11
list.ParagraphFormat.LineSpacing = 1.5 * 12
list.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple
section.Document.Styles.Add(list)
title = section.AddParagraph()
text = title.AppendText("Field type list:")
title.ApplyStyle(list.Name)
first = True
for fieldType in sorted(FieldType.__members__.values(),key=lambda c: c.value):
    if fieldType == FieldType.FieldUnknown or fieldType == FieldType.FieldNone or fieldType == FieldType.FieldEmpty:
        continue
    paragraph = section.AddParagraph()
    paragraph.AppendText("{0:s} is supported in Spire.Doc".format(fieldType.name))
    if first:
        paragraph.ListFormat.ApplyNumberedStyle()
        first = False
    else:
        paragraph.ListFormat.ContinueListNumbering()
    paragraph.ApplyStyle(list.Name)

#Save as doc file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
