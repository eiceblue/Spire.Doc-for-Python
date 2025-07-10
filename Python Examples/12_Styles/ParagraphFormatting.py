from spire.doc import *
from spire.doc.common import *


outputFile = "ParagraphFormatting.docx"
#Initialize a document
document = Document()
sec = document.AddSection()
para = sec.AddParagraph()
para.AppendText("Paragraph Formatting")
para.ApplyStyle(BuiltinStyle.Title)

para = sec.AddParagraph()
para.AppendText("This paragraph is surrounded with borders.")
para.Format.Borders.BorderType = BorderStyle.Single
para.Format.Borders.Color = Color.get_Red()

para = sec.AddParagraph()
para.AppendText("The alignment of this paragraph is Left.")
para.Format.HorizontalAlignment = HorizontalAlignment.Left

para = sec.AddParagraph()
para.AppendText("The alignment of this paragraph is Center.")
para.Format.HorizontalAlignment = HorizontalAlignment.Center

para = sec.AddParagraph()
para.AppendText("The alignment of this paragraph is Right.")
para.Format.HorizontalAlignment = HorizontalAlignment.Right

para = sec.AddParagraph()
para.AppendText("The alignment of this paragraph is justified.")
para.Format.HorizontalAlignment = HorizontalAlignment.Justify

para = sec.AddParagraph()
para.AppendText("The alignment of this paragraph is distributed.")
para.Format.HorizontalAlignment = HorizontalAlignment.Distribute

para = sec.AddParagraph()
para.AppendText("This paragraph has the gray shadow.")
para.Format.BackColor = Color.get_Gray()

para = sec.AddParagraph()
para.AppendText("This paragraph has the following indentations: Left indentation is 10pt, right indentation is 10pt, first line indentation is 15pt.")
para.Format.SetLeftIndent(10)
para.Format.SetRightIndent(10)
para.Format.SetFirstLineIndent(15)

para = sec.AddParagraph()
para.AppendText("The hanging indentation of this paragraph is 15pt.")
#Negative value represents hanging indentation
para.Format.SetFirstLineIndent(-15)

para = sec.AddParagraph()
para.AppendText("This paragraph has the following spacing: spacing before is 10pt, spacing after is 20pt, line spacing is at least 10pt.")
para.Format.AfterSpacing = 20
para.Format.BeforeSpacing = 10
para.Format.LineSpacingRule = LineSpacingRule.AtLeast
para.Format.LineSpacing = 10

#Save as docx file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()

